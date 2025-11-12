using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WIA; // Ensure COM reference to "Microsoft Windows Image Acquisition Library v2.0" is added

// This defines the FormatID constants locally, avoiding the interop embedding issue.
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
    class Program
    {
        // serialize scans so only one runs at a time
        static readonly SemaphoreSlim ScanSemaphore = new SemaphoreSlim(1, 1);
        static readonly TimeSpan ScanTimeout = TimeSpan.FromMinutes(5);

        const int WIA_DPS_DOCUMENT_HANDLING_SELECT = 3088;
        const int WIA_DPS_DOCUMENT_HANDLING_STATUS = 3089;
        const int WIA_DPS_PAGES = 3096;
        const int FEEDER = 1;
        const int FEED_READY = 1;
        const string wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";

        [STAThread]
        static void Main()
        {
            var listener = new HttpListener();
            string url = "http://localhost:9257/";
            listener.Prefixes.Add(url);
            listener.Start();
            Console.WriteLine($"Scanner agent running at {url}");
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
                if (req.Url.AbsolutePath == "/ping")
                {
                    WriteText(resp, "ok");
                    return;
                }

                if (req.Url.AbsolutePath == "/scan")
                {
                    HandleScan(resp);
                    return;
                }

                resp.StatusCode = 404;
                WriteText(resp, "not found");
            }
            catch (Exception ex)
            {
                // last-resort error handling for request processing
                WriteJsonArray(resp, new List<string>(), "Internal error: " + ex.Message);
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
            catch (Exception ex)
            {
                Console.WriteLine($"Could not get property {propertyId}: {ex.Message}");
            }
            return 0;
        }


        static void HandleScan(HttpListenerResponse resp)
        {
            // Do not allow concurrent scans
            if (!ScanSemaphore.Wait(0))
            {
                WriteJsonArray(resp, new List<string>(), "Another scan is in progress");
                return;
            }

            List<string> tempFiles = null;

            try
            {
                // Acquire images on dedicated STA thread via TaskCompletionSource
                var tcs = new TaskCompletionSource<(List<string> files, string status)>();

                Thread sta = new Thread(() =>
                {
                    CommonDialog dialog = null;
                    Device device = null;
                    var files = new List<string>();
                    string status = "ok";

                    try
                    {
                        dialog = new CommonDialog();

                        // Let user pick device (UI) -> returns Device or null
                        device = dialog.ShowSelectDevice(WiaDeviceType.ScannerDeviceType, false, false);
                        if (device == null)
                        {
                            tcs.SetResult((new List<string>(), "No scanner selected"));
                            return;
                        }

                        //try
                        //{
                        //    // Get the device's items (scanner)
                        //    Items items = device.Items;

                        //    // Configure common scanner settings
                        //    SetProperty(items, 6146, 200);  // Horizontal Resolution (DPI)
                        //    SetProperty(items, 6147, 200);  // Vertical Resolution (DPI)
                        //    SetProperty(items, 4104, 1);    // Color format: 1 = Color, 2 = Grayscale, 4 = B&W
                        //    SetProperty(items, 4106, 0);    // Brightness
                        //    SetProperty(items, 4107, 0);    // Contrast
                        //    SetProperty(items, 6154, 1);    // Page Size: 1 = Letter, etc.
                        //}
                        //catch (Exception ex)
                        //{
                        //    Console.WriteLine($"Warning: Could not set device properties: {ex.Message}");
                        //}


                        //// Acquire images: single ImageFile or collection (multi-page)
                        //object scanResult = dialog.ShowAcquireImage(
                        //    WiaDeviceType.ScannerDeviceType,
                        //    WiaImageIntent.UnspecifiedIntent,
                        //    WiaImageBias.MaximizeQuality,
                        //    "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}", // replace with literal value
                        //    true,  // show UI
                        //    true,  // allow multi-image (ADF)
                        //    false
                        //);


                        //if(scanResult != null)
                        //{
                        //    Type resultType = scanResult.GetType();
                        //    string typeName = resultType.Name;

                        //    Console.WriteLine($"Typename:  {typeName}");

                        //    if(typeName.Contains("vector") || typeName.Contains("Collection"))
                        //    {
                        //        Console.WriteLine("World - Multiple images detected");

                        //        try
                        //        {
                        //            // Use dynamic to access COM Vector
                        //            dynamic vector = scanResult;
                        //            int count = vector.Count;
                        //            Console.WriteLine($"Image count: {count}");

                        //            // WIA Vector is 1-indexed
                        //            for (int i = 1; i <= count; i++)
                        //            {
                        //                ImageFile img = (ImageFile)vector[i];
                        //                files.Add(SaveImageFileToTemp(img));
                        //                ReleaseCom(img);
                        //            }
                        //        }
                        //        catch (Exception ex)
                        //        {
                        //            Console.WriteLine($"Error processing vector: {ex.Message}");

                        //            // Fallback: try IEnumerable
                        //            if (scanResult is System.Collections.IEnumerable enumerable)
                        //            {
                        //                foreach (var item in enumerable)
                        //                {
                        //                    if (item == null) continue;
                        //                    ImageFile img = item as ImageFile ?? (ImageFile)item;
                        //                    files.Add(SaveImageFileToTemp(img));
                        //                    ReleaseCom(img);
                        //                }
                        //            }
                        //        }
                        //    }
                        //    else if (scanResult is ImageFile singleImage)
                        //    {
                        //        Console.WriteLine("Hello");
                        //        files.Add(SaveImageFileToTemp(singleImage));
                        //        ReleaseCom(singleImage);
                        //    }
                        //    //else if (scanResult is System.Collections.IEnumerable enumerable)
                        //    //{
                        //    //    Console.WriteLine("World");
                        //    //    foreach (var item in enumerable)
                        //    //    {
                        //    //        if (item == null) continue;
                        //    //        ImageFile img = item as ImageFile;
                        //    //        if (img == null)
                        //    //        {
                        //    //            try { img = (ImageFile)item; } catch { continue; }
                        //    //        }
                        //    //        files.Add(SaveImageFileToTemp(img));
                        //    //        ReleaseCom(img);
                        //    //    }
                        //    //}
                        //    else
                        //    {
                        //        status = "No image acquired";
                        //    }

                        //}
                        //  tcs.SetResult((files, status));

                        //**********************************************************************//
                        //Item item = device.Items[1]; // Get first item (scanner)

                        //// Configure ADF settings
                        //Program.SetProperty(item.Properties, "3088", 1); // Document Handling Select: 1 = Flatbed, 4 = ADF

                        //// Configure other properties
                        //Program.SetProperty(item.Properties, "6146", 200);  // Horizontal Resolution
                        //Program.SetProperty(item.Properties, "6147", 200);  // Vertical Resolution
                        //Program.SetProperty(item.Properties, "4104", 1);    // Color format: Color

                        //// Check if ADF is available and has documents
                        //int documentHandling = GetPropertyInt(item.Properties, "3088");
                        //bool hasADF = (documentHandling & 0x00000004) != 0; // Check ADF bit
                        //bool hasDocuments = (documentHandling & 0x00000002) != 0; // Check document feeder ready bit

                        //Console.WriteLine($"ADF available: {hasADF}, Documents ready: {hasDocuments}");

                        //if (hasADF && hasDocuments)
                        //{
                        //    // Scan multiple pages using ADF
                        //    bool morePages = true;
                        //    while (morePages)
                        //    {
                        //        try
                        //        {
                        //            ImageFile image = (ImageFile)item.Transfer(WiaFormatID.wiaFormatJPEG);
                        //            if (image != null)
                        //            {
                        //                files.Add(SaveImageFileToTemp(image));
                        //                ReleaseCom(image);
                        //                Console.WriteLine("Scanned page successfully");
                        //            }

                        //            // Check if more pages
                        //            documentHandling = GetPropertyInt(item.Properties, "3088");
                        //            morePages = (documentHandling & 0x00000003) != 0; // Check feeder has more documents
                        //        }
                        //        catch (Exception ex)
                        //        {
                        //            if (ex.Message.Contains("No documents") || ex.Message.Contains("feeder is empty"))
                        //            {
                        //                morePages = false;
                        //                Console.WriteLine("ADF empty");
                        //            }
                        //            else
                        //            {
                        //                throw;
                        //            }
                        //        }
                        //    }
                        //}
                        //else
                        //{
                        //    // Single page scan
                        //    ImageFile image = (ImageFile)item.Transfer(WiaFormatID.wiaFormatJPEG);
                        //    if (image != null)
                        //    {
                        //        files.Add(SaveImageFileToTemp(image));
                        //        ReleaseCom(image);
                        //    }
                        //}

                        SetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FEEDER);
                        // Set Pages to 0 for continuous scan until empty
                        SetProperty(device.Properties, WIA_DPS_PAGES, 0);

                        Console.WriteLine("KO");

                        bool hasMorePages = true;

                        while (hasMorePages)
                        {
                            try
                            {
                                // 3. Get the first item, which the driver uses as a proxy for the ADF stream
                                Item item = device.Items[1] as Item;

                                // 4. Transfer the image (this triggers the scan of one page)
                                //Console.WriteLine("pas");
                                //ImageFile imageFile = (ImageFile)item.Transfer(WiaFormatID.wiaFormatTIFF);
                                //Console.WriteLine("ad");
                                ImageFile imageFile = null;

                                try
                                {
                                    imageFile = (ImageFile)item.Transfer(WiaFormatID.wiaFormatTIFF);
                                    Console.WriteLine(imageFile.ToString());
                                }
                                catch (COMException transferEx)
                                {
                                    // Known WIA code when feeder empty: 0x80210003 (WINCODEC_ERR or WIA error); stop loop
                                    int hr = transferEx.ErrorCode;
                                    if (hr == unchecked((int)0x80210003))
                                    {
                                        Console.WriteLine("Feeder empty / no more pages.");
                                        break;
                                    }
                                    throw;
                                }

                                if (imageFile != null)
                                {
                                    files.Add(SaveImageFileToTemp(imageFile));
                                    ReleaseCom(imageFile);
                                }
                                
                                Console.WriteLine($"Scanned page {files.Count}");

                                // 5. Check if there are more pages ready in the feeder
                                int documentHandlingStatus = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_STATUS);
                                hasMorePages = ((documentHandlingStatus & FEED_READY) != 0);
                            }
                            catch (Exception ex)
                            {
                                // A common WIA error (0x80210003) indicates the feeder is empty.
                                // You may need to handle specific COM exceptions here.
                                if (ex.Message.Contains("0x80210003"))
                                {
                                    Console.WriteLine("ADF is empty. Scanning complete.");
                                    hasMorePages = false;
                                }
                                else
                                {
                                    Console.WriteLine($"An error occurred: {ex.Message}");
                                    hasMorePages = false; // Stop on unexpected errors
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        tcs.SetResult((new List<string>(), "Scan failed: " + ex.Message));
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
                    WriteJsonArray(resp, new List<string>(), "Scan timed out");
                    return;
                }

                var result = tcs.Task.Result;
                tempFiles = result.files;
                // By default return file paths to reduce memory. You can change this to return base64 by reading the files.
                WriteJsonArray(resp, result.files, result.status);
            }
            finally
            {

                // cleanup temp files
                if (tempFiles != null)
                {
                    foreach (var f in tempFiles)
                    {
                        try { File.Delete(f); } catch { }
                    }
                }

                ScanSemaphore.Release();
            }
        }
       

 
        // Save ImageFile to a temp file and return the path
        static string SaveImageFileToTemp(ImageFile img)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "scanneragent_" + Guid.NewGuid().ToString() + ".jpg");
            img.SaveFile(tempPath);
            return tempPath;
        }

        // Safe COM release helper
        static void ReleaseCom(object comObj)
        {
            if (comObj == null) return;
            try
            {
                while (Marshal.ReleaseComObject(comObj) > 0) { }
            }
            catch { /* ignore */ }
            finally
            {
                comObj = null;
            }
        }

        static void WriteText(HttpListenerResponse resp, string text)
        {
            resp.ContentType = "text/plain; charset=utf-8";
            var writer = new StreamWriter(resp.OutputStream);
            writer.Write(text);
            writer.Flush();
            resp.OutputStream.Close();
        }

        static void WriteJsonArray(HttpListenerResponse resp, List<string> imagesOrPaths, string status)
        {
            resp.ContentType = "application/json; charset=utf-8";
            resp.Headers.Add("Access-Control-Allow-Origin", "*"); // adjust for production

            // imagesOrPaths contains file paths (default). If you want base64, read files first.
            string itemsJson = string.Join(",", imagesOrPaths.ConvertAll(i => $"\"{EscapeJsonString(i)}\""));
            string json = $"{{\"status\":\"{EscapeJsonString(status)}\",\"files\":[{itemsJson}]}}";

            var writer = new StreamWriter(resp.OutputStream);
            writer.Write(json);
            writer.Flush();
            resp.OutputStream.Close();
        }

        static string EscapeJsonString(string s)
        {
            return s?.Replace("\\", "\\\\").Replace("\"", "\\\"") ?? string.Empty;
        }
    }
}