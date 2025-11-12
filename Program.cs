using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WIA;

static class WiaFormatID
{
    public const string wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
    public const string wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
    public const string wiaFormatGIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}";
    public const string wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
    public const string wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";
}

namespace ScannerAgent


{
    //class WIA_DPS_DOCUMENT_HANDLING_SELECT
    //{
    //    public const uint FEEDER = 0x00000001;
    //    public const uint FLATBED = 0x00000002;
    //}

    //class WIA_DPS_DOCUMENT_HANDLING_STATUS
    //{
    //    public const uint FEED_READY = 0x00000001;
    //}

    //class WIA_PROPERTIES
    //{
    //    public const uint WIA_RESERVED_FOR_NEW_PROPS = 1024;
    //    public const uint WIA_DIP_FIRST = 2;
    //    public const uint WIA_DPA_FIRST = WIA_DIP_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
    //    public const uint WIA_DPC_FIRST = WIA_DPA_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
    //    //
    //    // Scanner only device properties (DPS)
    //    //
    //    public const uint WIA_DPS_FIRST = WIA_DPC_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
    //    public const uint WIA_DPS_DOCUMENT_HANDLING_STATUS = WIA_DPS_FIRST + 13;
    //    public const uint WIA_DPS_DOCUMENT_HANDLING_SELECT = WIA_DPS_FIRST + 14;
    //}

    class WIA_ERRORS
    {
        public const uint BASE_VAL_WIA_ERROR = 0x80210000;
        public const uint WIA_ERROR_PAPER_EMPTY = BASE_VAL_WIA_ERROR + 3;
    }
    class Program
    {
        static readonly SemaphoreSlim ScanSemaphore = new SemaphoreSlim(1, 1);
        static readonly TimeSpan ScanTimeout = TimeSpan.FromMinutes(5);

        const int WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES = 3086;
        const int WIA_DPS_DOCUMENT_HANDLING_SELECT = 3088;
        const int WIA_DPS_DOCUMENT_HANDLING_STATUS = 3087;
        const int WIA_DPS_PAGES = 3098;
        const int WIA_IPS_XRES = 6147;
        const int WIA_IPS_YRES = 6148;
        const int WIA_IPS_CUR_INTENT = 6146;

        const int FEEDER = 1;
        const int FLATBED = 4;
        const int DUPLEX = 8;
        const int FEED_READY = 1;

        [STAThread]
        static void Main()
        {
            var listener = new HttpListener();
            string url = "http://localhost:9257/";
            listener.Prefixes.Add(url);
            listener.Start();
            Console.WriteLine($"Scanner agent running at {url}");
            Console.WriteLine("Endpoints:");
            Console.WriteLine("  GET  /ping                  - Health check");
            Console.WriteLine("  POST /scan?mode=flatbed     - Scan single page from flatbed");
            Console.WriteLine("  POST /scan?mode=adf         - Scan multiple pages from ADF");
            Console.WriteLine("  POST /scan?mode=auto        - Auto-detect and use available mode");
            Console.WriteLine("  POST /scan?mode=X&format=Y  - Specify format (jpeg/png/tiff)");
            Console.WriteLine("Press Ctrl+C to stop.");

            while (true)
            {
                var context = listener.GetContext();
                Task.Run(() => ProcessRequest(context));
            }
        }

        static void ProcessRequest(HttpListenerContext context)
        {
            var req = context.Request;
            var resp = context.Response;

            try
            {
                // Handle CORS preflight
                if (req.HttpMethod == "OPTIONS")
                {
                    resp.Headers.Add("Access-Control-Allow-Origin", "*");
                    resp.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
                    resp.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
                    resp.StatusCode = 200;
                    resp.Close();
                    return;
                }

                if (req.Url.AbsolutePath == "/ping")
                {
                    WriteText(resp, "ok");
                    return;
                }

                if (req.Url.AbsolutePath == "/scan")
                {
                    string mode = req.QueryString["mode"] ?? "auto";
                    string format = req.QueryString["format"] ?? "jpeg";
                    int dpi = int.Parse(req.QueryString["dpi"] ?? "300");
                    HandleScan(resp, mode, format, dpi);
                    return;
                }

                resp.StatusCode = 404;
                WriteText(resp, "not found");
            }
            catch (Exception ex)
            {
                WriteJsonResponse(resp, new List<ScannedImage>(), "Internal error: " + ex.Message);
            }
        }

        private static void SetProperty(Properties properties, int propertyId, object value)
        {
            try
            {
                foreach (Property prop in properties)
                {
                    if (prop.PropertyID == propertyId)
                    {
                        prop.set_Value(value);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not set property {propertyId}: {ex.Message}");
            }
        }

        private static int GetProperty(Properties properties, int propertyId)
        {
            try
            {
                foreach (Property prop in properties)
                {
                    
                    if (prop.PropertyID == propertyId)
                    {
                        return Convert.ToInt32(prop.get_Value());
                    }
                }
            }
            catch { }
            return 0;
        }

        static void HandleScan(HttpListenerResponse resp, string mode, string format, int dpi)
        {
            if (!ScanSemaphore.Wait(0))
            {
                WriteJsonResponse(resp, new List<ScannedImage>(), "Another scan is in progress");
                return;
            }

            try
            {
                var tcs = new TaskCompletionSource<(List<ScannedImage> images, string status)>();

                Thread sta = new Thread(() =>
                {
                    CommonDialog dialog = null;
                    Device device = null;
                    var images = new List<ScannedImage>();
                    string status = "ok";

                    try
                    {
                        dialog = new CommonDialog();
                        device = dialog.ShowSelectDevice(WiaDeviceType.ScannerDeviceType, false, false);

                        if (device == null)
                        {
                            tcs.SetResult((new List<ScannedImage>(), "No scanner selected"));
                            return;
                        }

                        Console.WriteLine($"Selected device: {device.DeviceID}");

                        int capabilities = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES);
                        bool hasFlatbed = (capabilities & FLATBED) != 0;
                        bool hasFeeder = (capabilities & FEEDER) != 0;

                        Console.WriteLine($"Capabilities: {capabilities}");

                        Console.WriteLine($"Flatbed: {hasFlatbed}, Feeder: {hasFeeder}");

                        string actualMode = mode.ToLower();
                        if (actualMode == "auto")
                        {
                            if (hasFeeder)
                            {
                                int handlingStatus = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_STATUS);
                                bool feederReady = (handlingStatus & FEED_READY) != 0;
                                actualMode = feederReady ? "adf" : "flatbed";
                            }
                            else
                            {
                                actualMode = "flatbed";
                            }
                            Console.WriteLine($"Auto mode selected: {actualMode}");
                        }

                        
                        string wiaFormat = GetWiaFormat(format);

                        if (actualMode == "adf" && hasFeeder)
                        {
                            images = ScanFromFeeder(device, wiaFormat, dpi);
                        }
                        else if (actualMode == "flatbed" || !hasFeeder)
                        {
                            images = ScanFromFlatbed(device, wiaFormat, dpi);
                        }
                        else
                        {
                            status = $"Requested mode '{mode}' not available";
                        }

                        tcs.SetResult((images, status));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Scan error: {ex.Message}");
                        tcs.SetResult((new List<ScannedImage>(), "Scan failed: " + ex.Message));
                    }
                    finally
                    {
                        ReleaseCom(device);
                        ReleaseCom(dialog);
                    }
                });

                sta.SetApartmentState(ApartmentState.STA);
                sta.IsBackground = true;
                sta.Start();

                if (!tcs.Task.Wait(ScanTimeout))
                {
                    WriteJsonResponse(resp, new List<ScannedImage>(), "Scan timed out");
                    return;
                }

                var result = tcs.Task.Result;
                Console.WriteLine($"Scan completed: {result.images.Count} pages");
                WriteJsonResponse(resp, result.images, result.status);
            }
            finally
            {
                ScanSemaphore.Release();
            }
        }

        static string GetWiaFormat(string format)
        {
            switch (format.ToLower())
            {
                case "png": return WiaFormatID.wiaFormatPNG;
                case "tiff": return WiaFormatID.wiaFormatTIFF;
                case "bmp": return WiaFormatID.wiaFormatBMP;
                default: return WiaFormatID.wiaFormatJPEG;
            }
        }

        static List<ScannedImage> ScanFromFlatbed(Device device, string wiaFormat, int dpi)
        {
            Console.WriteLine("Scanning from flatbed...");
            var images = new List<ScannedImage>();

            try
            {
                SetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FLATBED);

                Item item = device.Items[1] as Item;

                if (item == null)
                {
                   throw new Exception("No flatbed item found");
                }

                Console.WriteLine($"Item name: {item.ItemID}");

                SetProperty(item.Properties, WIA_IPS_XRES, dpi);
                SetProperty(item.Properties, WIA_IPS_YRES, dpi);

                try
                {
                    SetProperty(item.Properties, WIA_IPS_CUR_INTENT, 4); // Color
                }
                catch (Exception intentEx)
                {
                    Console.WriteLine($"Warning: Could not set color intent: {intentEx.Message}");
                }


                Console.WriteLine("Starting transfer...");

                ImageFile imageFile = null;
                try
                {
                    imageFile = (ImageFile)item.Transfer(wiaFormat);
                }
                catch (COMException comEx)
                {
                    int errorCode = comEx.ErrorCode;
                    Console.WriteLine($"COM Exception during transfer: 0x{errorCode:X8}");
                    Console.WriteLine($"Error message: {comEx.Message}");

                    // Common WIA error codes
                    switch (errorCode)
                    {
                        case unchecked((int)0x80210015): // WIA_ERROR_PAPER_EMPTY
                            throw new Exception("No document in scanner");
                        case unchecked((int)0x80210006): // WIA_ERROR_PAPER_JAM
                            throw new Exception("Paper jam detected");
                        case unchecked((int)0x80210001): // WIA_ERROR_GENERAL_ERROR
                            throw new Exception("General scanner error - check if scanner is ready");
                        case unchecked((int)0x8021000C): // WIA_ERROR_DEVICE_BUSY
                            throw new Exception("Scanner is busy");
                        case unchecked((int)0x80210005): // WIA_ERROR_OFFLINE
                            throw new Exception("Scanner is offline");
                        default:
                            throw new Exception($"Scanner error: 0x{errorCode:X8} - {comEx.Message}");
                    }
                }

                if (imageFile != null)
                {
                    images.Add(ConvertImageToBase64(imageFile, 1));
                    ReleaseCom(imageFile);
                    Console.WriteLine("Flatbed scan completed");
                }
            }
            catch (COMException comEx)
            {
                Console.WriteLine($"COM Exception: 0x{comEx.ErrorCode:X8} - {comEx.Message}");
                throw new Exception($"Scanner COM error: 0x{comEx.ErrorCode:X8} - {comEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Flatbed scan error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }

            return images;
        }

        static List<ScannedImage> ScanFromFeeder(Device device, string wiaFormat, int dpi)
        {
            Console.WriteLine("Scanning from ADF...");
            var images = new List<ScannedImage>();

            try
            {
                SetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FEEDER);
                //SetProperty(device.Properties, WIA_DPS_PAGES, 0);
                    
                bool hasMorePages = true;
                int pageCount = 0;
                int x = 0;

                while (hasMorePages)
                {
                    DeviceManager manager = (DeviceManager)Activator.CreateInstance(Type.GetTypeFromProgID("WIA.DeviceManager"));
                    Device WiaDev = null;
                    foreach (DeviceInfo info in manager.DeviceInfos)
                    {
                        if (info.DeviceID == device.DeviceID)
                        {
                            Properties infoprop = null;
                            infoprop = info.Properties;

                            //connect to scanner
                            WiaDev = info.Connect();


                            break;
                        }
                    }


                    ImageFile imageFile = null;
                    Item item = device.Items[1] as Item;
                    try
                    {
                       
                        SetProperty(item.Properties, WIA_IPS_XRES, dpi);
                        SetProperty(item.Properties, WIA_IPS_YRES, dpi);
                        SetProperty(item.Properties, WIA_IPS_CUR_INTENT, 4);

                        try
                        {
                            imageFile = (ImageFile)item.Transfer(wiaFormat);
                        }
                        catch (COMException comEx)
                        {
                            int hr = comEx.ErrorCode;
                            if (hr == unchecked((int)0x80210003) || hr == unchecked((int)0x80210006))
                            {
                                Console.WriteLine("Feeder empty");
                                break;
                            }
                            throw;
                        }

                        if (imageFile != null)
                        {
                            pageCount++;
                            images.Add(ConvertImageToBase64(imageFile, pageCount));
                            ReleaseCom(imageFile);
                            Console.WriteLine($"Scanned page {pageCount}");
                            Console.WriteLine("OKA");
                        }

             
                    }
                   
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                        hasMorePages = false;
                    }

                    finally
                    {
                        item = null;
                        //determine if there are any more pages waiting

                        Property documentHandlingSelect = null;
                        Property documentHandlingStatus = null;
                        foreach (Property prop in WiaDev.Properties)
                        {
                            Console.WriteLine("Prop Id: " + prop.PropertyID+ " WIA DPS HAND SELECT: "+WIA_DPS_DOCUMENT_HANDLING_SELECT);
                            if (prop.PropertyID == WIA_DPS_DOCUMENT_HANDLING_SELECT)
                                documentHandlingSelect = prop;

                            Console.WriteLine("WIA DPS STATUS: " + WIA_DPS_DOCUMENT_HANDLING_STATUS);
                            if (prop.PropertyID == WIA_DPS_DOCUMENT_HANDLING_STATUS)
                                documentHandlingStatus = prop;


                        }

                        hasMorePages = false; //assume there are no more pages
                        if (documentHandlingSelect != null)
                        //may not exist on flatbed scanner but required for feeder
                        {
                            //check for document feeder
                            if ((Convert.ToUInt32(documentHandlingSelect.get_Value()) & FEEDER) != 0)
                            {
                                hasMorePages = ((Convert.ToUInt32(documentHandlingStatus.get_Value()) & FEED_READY) != 0);
                            }
                        }
                        x++;
                        Console.WriteLine("Loop " + x.ToString());
                        Console.WriteLine(hasMorePages.ToString());
                    }
                }

                Console.WriteLine($"ADF completed: {pageCount} pages");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Feeder error: {ex.Message}");
                throw;
            }

            return images;
        }

        static ScannedImage ConvertImageToBase64(ImageFile imageFile, int pageNumber)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".tmp");

            try
            {
                imageFile.SaveFile(tempPath);
                byte[] imageBytes = File.ReadAllBytes(tempPath);
                string base64 = Convert.ToBase64String(imageBytes);

                string mimeType = "image/jpeg";
                if (imageFile.FormatID == WiaFormatID.wiaFormatPNG) mimeType = "image/png";
                else if (imageFile.FormatID == WiaFormatID.wiaFormatTIFF) mimeType = "image/tiff";

                return new ScannedImage
                {
                    PageNumber = pageNumber,
                    Base64Data = $"data:{mimeType};base64,{base64}",
                    Size = imageBytes.Length,
                    Format = mimeType
                };
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }

        static void ReleaseCom(object comObj)
        {
            if (comObj == null) return;
            try
            {
                while (Marshal.ReleaseComObject(comObj) > 0) { }
            }
            catch { }
        }

        static void WriteText(HttpListenerResponse resp, string text)
        {
            resp.ContentType = "text/plain; charset=utf-8";
            resp.Headers.Add("Access-Control-Allow-Origin", "*");
            var writer = new StreamWriter(resp.OutputStream);
            writer.Write(text);
            writer.Flush();
            resp.OutputStream.Close();
        }

        static void WriteJsonResponse(HttpListenerResponse resp, List<ScannedImage> images, string status)
        {
            resp.ContentType = "application/json; charset=utf-8";
            resp.Headers.Add("Access-Control-Allow-Origin", "*");

            var imagesJson = new List<string>();
            foreach (var img in images)
            {
                imagesJson.Add($"{{\"pageNumber\":{img.PageNumber},\"base64\":\"{EscapeJsonString(img.Base64Data)}\",\"size\":{img.Size},\"format\":\"{img.Format}\"}}");
            }

            string json = $"{{\"status\":\"{EscapeJsonString(status)}\",\"pageCount\":{images.Count},\"images\":[{string.Join(",", imagesJson)}]}}";

            var writer = new StreamWriter(resp.OutputStream);
            writer.Write(json);
            writer.Flush();
            resp.OutputStream.Close();
        }

        static string EscapeJsonString(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
        }
    }

    class ScannedImage
    {
        public int PageNumber { get; set; }
        public string Base64Data { get; set; }
        public long Size { get; set; }
        public string Format { get; set; }
    }
}