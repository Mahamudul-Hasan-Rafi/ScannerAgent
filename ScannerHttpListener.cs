using EScanner.Controllers;
using EScanner.Model;
using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace EScanner
{
    public class ScannerHttpListener
    {
        private readonly string _url;
        private readonly HttpListener _listener;
        private readonly ScannerController _controller;

        public ScannerHttpListener(string url, ScannerController controller)
        {
            _url = url;
            _listener = new HttpListener();
            _listener.Prefixes.Add(_url);
            _controller = controller;
        }

        public void Start()
        {
            _listener.Start();
            Console.WriteLine($"Listening on {_url}");
            Task.Run(() => ListenLoop());
        }

        private async Task ListenLoop()
        {
            while (true)
            {
                var context = await _listener.GetContextAsync();
                _ = Task.Run(() => ProcessRequest(context));
            }
        }

        private async Task ProcessRequest(HttpListenerContext context)
        {
            var req = context.Request;
            var resp = context.Response;

            try
            {
                // Handle CORS
                if (req.HttpMethod == "OPTIONS")
                {
                    resp.Headers.Add("Access-Control-Allow-Origin", "*");
                    resp.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
                    resp.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
                    resp.StatusCode = 200;
                    resp.Close();
                    return;
                }

                // Ping
                if (req.Url.AbsolutePath == "/ping" && req.HttpMethod == "GET")
                {
                    WriteText(resp, "ok");
                    return;
                }

                // Scan
                if (req.Url.AbsolutePath == "/scan" && req.HttpMethod == "POST")
                {
                    string requestBody;
                    using (var reader = new StreamReader(req.InputStream))
                        requestBody = reader.ReadToEnd();

                    var scanRequest = JsonConvert.DeserializeObject<ScanRequest>(requestBody);
                    var result = await _controller.ScanAsync(scanRequest);

                    WriteJsonResponse(resp, result);
                    return;
                }

                // Devices
                if (req.Url.AbsolutePath == "/devices" && req.HttpMethod == "GET")
                {
                    var devices = _controller.Devices();
                    WriteJsonResponse(resp, devices);
                    return;
                }

                // Not found
                resp.StatusCode = 404;
                WriteJsonResponse(resp, new { error = "Endpoint not found" });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                resp.StatusCode = 500;
                WriteJsonResponse(resp, new
                {
                    error = "Internal server error",
                    message = ex.Message
                });
            }
        }

        private void WriteText(HttpListenerResponse resp, string text)
        {
            resp.ContentType = "text/plain; charset=utf-8";
            resp.Headers.Add("Access-Control-Allow-Origin", "*");
            var writer = new StreamWriter(resp.OutputStream);
            writer.Write(text);
            writer.Flush();
            resp.OutputStream.Close();
        }

        private void WriteJsonResponse(HttpListenerResponse resp, object obj)
        {
            try
            {
                string json = JsonConvert.SerializeObject(obj);
                resp.ContentType = "application/json; charset=utf-8";
                resp.Headers.Add("Access-Control-Allow-Origin", "*");
                using (var writer = new StreamWriter(resp.OutputStream))
                {
                    writer.Write(json);
                    writer.Flush();
                }
            }
            catch
            {
                // In case even writing fails, don't throw
            }
            finally
            {
                try { resp.OutputStream.Close(); } catch { }
            }
        }
    }
}
