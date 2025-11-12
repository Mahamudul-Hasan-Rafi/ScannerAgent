using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NTwain;
using NTwain.Data;

class Program
{
    enum RequestType { ListSources, StartScan }

    class TwainRequest
    {
        public RequestType Type;
        public TaskCompletionSource<object> Tcs = new TaskCompletionSource<object>();
        public int SelectedIndex;
        public int TimeoutMs = 60000;
    }

    static TwainSession twain;
    static DataSource currentSource;
    static List<string> scannedBase64;
    static HttpListener listener;
    static BlockingCollection<TwainRequest> requestQueue = new BlockingCollection<TwainRequest>();

    [STAThread]
    static void Main()
    {
        // Start TWAIN on dedicated STA thread (required for many drivers)
        var twainThread = new Thread(TwainThreadProc) { IsBackground = true };
        twainThread.SetApartmentState(ApartmentState.STA);
        twainThread.Start();

        // Start HTTP listener on main thread (console)
        listener = new HttpListener();
        listener.Prefixes.Add("http://localhost:9257/");
        listener.Start();
        Console.WriteLine("Scanner API running at http://localhost:9257/");

        while (true)
        {
            var context = listener.GetContext();
            var req = context.Request;
            var resp = context.Response;

            if (req.Url.AbsolutePath == "/ping")
            {
                WriteText(resp, "ok");
            }
            else if (req.Url.AbsolutePath == "/scan")
            {
                HandleScan(resp);
            }
            else
            {
                resp.StatusCode = 404;
                WriteText(resp, "not found");
            }
        }
    }

    // STA thread proc that owns TWAIN session and driver dialogs
    static void TwainThreadProc()
    {
        twain = new TwainSession(TWIdentity.CreateFromAssembly(DataGroups.Image, typeof(Program).Assembly));
        twain.TransferReady += Twain_TransferReady;
        twain.DataTransferred += Twain_DataTransferred;
        twain.SourceDisabled += Twain_SourceDisabled;
        twain.Open();

        foreach (var request in requestQueue.GetConsumingEnumerable())
        {
            try
            {
                if (request.Type == RequestType.ListSources)
                {
                    var sources = twain.GetSources();
                    var names = new List<string>();
                    foreach (var s in sources)
                        names.Add(s.Name ?? $"Scanner {s.Id}");
                    request.Tcs.TrySetResult(names);
                }
                else if (request.Type == RequestType.StartScan)
                {
                    scannedBase64 = new List<string>();
                    var sources = new List<DataSource>(twain.GetSources());
                    if (sources.Count == 0)
                    {
                        request.Tcs.TrySetResult(scannedBase64);
                        continue;
                    }

                    int idx = Math.Max(0, Math.Min(request.SelectedIndex, sources.Count - 1));
                    currentSource = sources[idx];
                    currentSource.Open();
                    currentSource.Enable(SourceEnableMode.ShowUI, false, IntPtr.Zero);

                    // Pump messages while driver UI is open
                    int waited = 0;
                    while (currentSource != null && waited < request.TimeoutMs)
                    {
                        Application.DoEvents();
                        Thread.Sleep(100);
                        waited += 100;
                    }

                    request.Tcs.TrySetResult(scannedBase64 ?? new List<string>());
                }
            }
            catch (Exception ex)
            {
                request.Tcs.TrySetException(ex);
                currentSource = null;
            }
        }
    }

    static void HandleScan(HttpListenerResponse resp)
    {
        // 1) Ask TWAIN thread for sources
        var listReq = new TwainRequest { Type = RequestType.ListSources };
        requestQueue.Add(listReq);
        if (!listReq.Tcs.Task.Wait(5000))
        {
            WriteJson(resp, new List<string>(), "Failed to enumerate scanners");
            return;
        }

        var names = (List<string>)listReq.Tcs.Task.Result;
        if (names.Count == 0)
        {
            WriteJson(resp, new List<string>(), "No scanner found");
            return;
        }

        // 2) Let user pick one via console (console UI)
        Console.WriteLine("Select a scanner:");
        for (int i = 0; i < names.Count; i++)
            Console.WriteLine($"{i + 1}: {names[i]}");

        int choice = 1;
        try
        {
            Console.Write("Enter scanner number: ");
            choice = int.Parse(Console.ReadLine());
        }
        catch { choice = 1; }

        // 3) Send scan request with chosen index
        var scanReq = new TwainRequest { Type = RequestType.StartScan, SelectedIndex = choice - 1 };
        requestQueue.Add(scanReq);

        if (!scanReq.Tcs.Task.Wait(scanReq.TimeoutMs + 1000))
        {
            WriteJson(resp, new List<string>(), "Scan timed out");
            return;
        }

        if (scanReq.Tcs.Task.IsFaulted)
        {
            var ex = scanReq.Tcs.Task.Exception?.GetBaseException();
            WriteJson(resp, new List<string>(), "Scan failed: " + (ex?.Message ?? "unknown"));
            return;
        }

        var images = (List<string>)scanReq.Tcs.Task.Result;
        WriteJson(resp, images, "ok");
    }

    private static void Twain_DataTransferred(object sender, DataTransferredEventArgs e)
    {
        if (e.NativeData != IntPtr.Zero)
        {
            using (var img = System.Drawing.Image.FromHbitmap(e.NativeData))
            using (var ms = new MemoryStream())
            {
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                string base64 = Convert.ToBase64String(ms.ToArray());
                scannedBase64.Add(base64);
                Console.WriteLine($"Captured page {scannedBase64.Count}");
            }
        }
    }

    private static void Twain_TransferReady(object sender, EventArgs e)
    {
        Console.WriteLine("Transfer ready...");
    }

    private static void Twain_SourceDisabled(object sender, EventArgs e)
    {
        if (scannedBase64 == null) scannedBase64 = new List<string>();

        if (scannedBase64.Count == 0)
            Console.WriteLine("Scan canceled by user.");
        else
            Console.WriteLine($"Scan finished! Total pages: {scannedBase64.Count}");

        try { currentSource?.Close(); } catch { }
        currentSource = null;
    }

    static void WriteText(HttpListenerResponse resp, string text)
    {
        using (var writer = new StreamWriter(resp.OutputStream))
        {
            writer.Write(text);
        }
    }

    static void WriteJson(HttpListenerResponse resp, List<string> images, string status)
    {
        resp.ContentType = "application/json";
        resp.Headers.Add("Access-Control-Allow-Origin", "*");

        string imagesJson = string.Join(",", images.ConvertAll(img => $"\"{img}\""));
        string json = $"{{\"status\":\"{status}\",\"images\":[{imagesJson}]}}";

        using (var writer = new StreamWriter(resp.OutputStream))
        {
            writer.Write(json);
        }
    }

}


// Install-Package Command: Install-Package NTwain -ProjectName ScannerAgent

// Data Flow - Web API → Twain → Scanner → Base64

/***
 * 
 * On /scan:

    Lists scanners → user selects one.

    Opens scanner UI → user picks ADF/flatbed, resolution, etc.

    Scans all pages asynchronously.

    Converts each page to Base64.

    Waits until scan is done.

    Returns JSON array of Base64 strings.

**/





// Further Enhancement

/***
 
Showing a proper Windows GUI for scanner/source selection

Returning results reliably for web API calls

Handling errors and timeouts

a nicer UI, host a WinForms/WPF window on the STA thread instead of console prompts.

**/