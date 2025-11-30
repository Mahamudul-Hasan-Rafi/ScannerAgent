using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ScannerAgent.Model;
using ScannerAgent.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WIA;
using System.Diagnostics;

namespace ScannerAgent.Services
{
    public class ScannerService
    {
        private readonly SemaphoreSlim _scanSemaphore = new SemaphoreSlim(1, 1);
        private readonly TimeSpan _scanTimeout = TimeSpan.FromMinutes(5); // up to 20 ADF pages at 15s each

        const int WIA_DPS_DOCUMENT_HANDLING_SELECT = 3088;   
        const int WIA_DPS_DOCUMENT_HANDLING_STATUS = 3087;

        const int WIA_IPS_XRES = 6147;
        const int WIA_IPS_YRES = 6148;
        const int WIA_IPS_CUR_INTENT = 6146;

        const int FEEDER = 1;
        const int FLATBED = 4;
        const int FEED_READY = 1;

        public async Task<(List<ScannedImage> images, string status, int statusCode)> ScanAsync(ScanRequest request)
        {
            if (!_scanSemaphore.Wait(0))
                return (new List<ScannedImage>(), "Another scan is in progress", 423);

            try
            {
                var tcs = new TaskCompletionSource<(List<ScannedImage>, string, int)>();

                Thread staThread = new Thread(() =>
                {
                    Device device = null;
                    var images = new List<ScannedImage>();

                    try
                    {
                        // Connect to device
                        device = ConnectToDeviceById(request.DeviceId);
                        if (device == null)
                        {
                            tcs.SetResult((new List<ScannedImage>(), $"Device '{request.DeviceId}' not found", 404));
                            return;
                        }

                        // ---------------------------
                        //  FIXED FEEDER/FLATBED DETECTION
                        // ---------------------------
                        int selectFlags = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT);

                        bool hasFeeder = (selectFlags & FEEDER) != 0;   
                        bool hasFlatbed = (selectFlags & FLATBED) != 0; 

                        string mode = request.Mode?.ToLower() ?? "auto";

                        if (mode == "auto")
                        {
                            if (hasFeeder)
                            {
                                int status = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_STATUS);

                                // If feeder exists AND has paper → use feeder
                                mode = (status & FEED_READY) != 0 ? "adf" : "flatbed";
                            }
                            else
                            {
                                mode = "flatbed";
                            }
                        }

                        string wiaFormat = ScannerUtils.GetWiaFormat(request.Format ?? "jpeg");
                        int dpi = request.Dpi ?? 300;
                        int color = request.Color ?? 4;

                        if (mode == "adf" && hasFeeder)
                            images = ScanFromFeeder(device, wiaFormat, dpi, color);
                        else
                            images = ScanFromFlatbed(device, wiaFormat, dpi, color);

                        tcs.SetResult((images, "ok", 200));
                    }
                    catch (Exception ex)
                    {
                        tcs.SetResult((new List<ScannedImage>(), "Scan failed: " + ex.Message, 500));
                    }
                    finally
                    {
                        ScannerUtils.ReleaseCom(device);
                    }
                });

                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();

                if (!tcs.Task.Wait(_scanTimeout))
                    return (new List<ScannedImage>(), "Scan timed out", 408);

                return await tcs.Task;  // tcs.Task.Result;
            }
            finally
            {
                _scanSemaphore.Release();
            }
        }


        public async Task StreamScanAsync(HttpListenerContext context, ScanRequest request)
        {
            if (!_scanSemaphore.Wait(0))
            {
                WriteStreamText(context.Response.OutputStream, "{\"status\":\"Another scan is in progress\", \"statusCode\":423}\n");
                try
                {
                    context.Response.OutputStream.Flush();
                    context.Response.Close();
                }
                catch
                {
                    try { context.Response.OutputStream.Close(); } catch { }
                }
                return;
            }

            var resp = context.Response;
            var output = resp.OutputStream;

            try
            {
                resp.StatusCode = 200;
                resp.ContentType = "application/x-ndjson"; // streaming JSON
                resp.SendChunked = true;


                var tcs = new TaskCompletionSource<bool>();

                Thread staThread = new Thread(() =>
                {
                    Device device = null;

                    try
                    {
                        // Connect device
                        device = ConnectToDeviceById(request.DeviceId);
                        if (device == null)
                        {
                            WriteStreamText(output, "{\"status\":\"Device not found\", \"statusCode\":404}\n");
                            tcs.SetResult(true);
                            return;
                        }

                        // ---------------------------
                        // MODE DETECTION
                        // ---------------------------
                        int selectFlags = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT);
                        bool hasFeeder = (selectFlags & FEEDER) != 0;
                        bool hasFlatbed = (selectFlags & FLATBED) != 0;

                        string mode = request.Mode?.ToLower() ?? "auto";

                        if (mode == "auto")
                        {
                            if (hasFeeder)
                            {
                                int status = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_STATUS);
                                mode = (status & FEED_READY) != 0 ? "adf" : "flatbed";
                            }
                            else mode = "flatbed";
                        }

                        // SELECT MODE
                        if (mode == "adf" && hasFeeder)
                            SetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FEEDER);
                        else
                            SetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FLATBED);

                        // SETTINGS
                        string wiaFormat = ScannerUtils.GetWiaFormat(request.Format ?? "jpeg");
                        int dpi = request.Dpi ?? 300;
                        int color = request.Color ?? 4;

                        // determine whether the device is currently using the feeder (ADF)
                        bool feederSelected = (GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT) & FEEDER) != 0;

                        // ---------------------------
                        // STREAM PAGE-BY-PAGE
                        // ---------------------------
                        foreach (var scanned in StreamPages(device, wiaFormat, dpi, color, feederSelected))
                        {
                            var j = JObject.FromObject(scanned);
                            j["status"] = "ok";
                            j["statusCode"] = 200;
                            WriteStreamText(output, j.ToString(Formatting.None) + "\n");
                        }

                        // End message
                        // followed by the thread returning, which causes the output stream to close.
                        WriteStreamText(output, "{\"status\":\"done\"}\n");
                        tcs.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        WriteStreamText(output, "{\"status\":\"error\", \"message\":\"" +
                            ex.Message.Replace("\"", "'") + "\"}\n");
                        tcs.SetResult(true);
                    }
                    finally
                    {
                        ScannerUtils.ReleaseCom(device);
                    }
                });

                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();

                // Wait for scan completion or overall timeout
                var finished = await Task.WhenAny(tcs.Task, Task.Delay(_scanTimeout));

                if (finished != tcs.Task)
                {
                    // worker didn't finish in time — notify client
                    WriteStreamText(output, "{\"status\":\"timeout\",\"statusCode\":408}\n");
                }
                else
                {
                    // ensure worker completed
                    await tcs.Task;
                }

                // Close the response so clients reliably observe stream end
                try
                {
                    resp.OutputStream.Flush();
                    resp.Close();
                }
                catch
                {
                    try { resp.OutputStream.Close(); } catch { }
                }

            }
            catch (Exception ex)
            {
                // best-effort write to stream if something unexpected escaped
                try
                {
                    WriteStreamText(context.Response.OutputStream, "{\"status\":\"error\", \"message\":\"" +
                        ex.Message.Replace("\"", "'") + "\"}\n");
   
                } catch { }
                throw;
            }
            finally
            {
                // safe finalization: flush and close the response so clients observe EOF
                try
                {
                    output.Flush();
                    resp.Close();
                }
                catch
                {
                    try { output.Close(); } catch { }
                }

                _scanSemaphore.Release();
            }
        }





        public DeviceListResponse ListDevices()
        {

            var response = new DeviceListResponse
            {
                devices = new List<DeviceInformation>(),
                status = "ok",
                statusCode = 200
            };

            //var devicesList = new List<DeviceInformation>();

            try
            {
                var manager = new DeviceManager();

                foreach (DeviceInfo info in manager.DeviceInfos)
                {
                    string id = info.DeviceID ?? "";
                    string name = null;
                    string type = info.Type.ToString();

                    try
                    {
                        var props = info.Properties;

                        try
                        {
                            var n = props["Name"];
                            if (n != null)
                                name = n.get_Value() as string;
                        }
                        catch { }

                        if (string.IsNullOrEmpty(name))
                        {
                            foreach (Property prop in props)
                            {
                                if (prop.PropertyID == 2)
                                {
                                    name = prop.get_Value()?.ToString();
                                    break;
                                }
                            }
                        }
                    }
                    catch { }

                    if (string.IsNullOrEmpty(name))
                        name = id;

                    response.devices.Add(new DeviceInformation
                    {
                        id = id,
                        name = name,
                        type = type
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing devices: {ex.Message}");
                // return an error response instead of swallowing the exception
                response.devices = new List<DeviceInformation>();
                response.status = "Error listing devices: " + ex.Message;
                response.statusCode = 500;
            }

            return response;
        }


        private Device ConnectToDeviceById(string deviceId)
        {
            deviceId = deviceId.Replace("\\\\", "\\");
            if (string.IsNullOrEmpty(deviceId)) return null;

            var manager = new DeviceManager();
            foreach (DeviceInfo info in manager.DeviceInfos)
            {
                if (string.Equals(info.DeviceID, deviceId, StringComparison.OrdinalIgnoreCase))
                {
                    try { return info.Connect(); }
                    catch { return null; }
                }
            }
            return null;
        }

        // --------------------------------------------------------------------
        // FLATBED
        // --------------------------------------------------------------------
        private List<ScannedImage> ScanFromFlatbed(Device device, string wiaFormat, int dpi, int color)
        {
            var images = new List<ScannedImage>();

            // 3088 = SELECT → force flatbed
            SetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FLATBED);

            var item = device.Items[1] as Item;
            if (item == null) throw new Exception("No flatbed item found");

            SetProperty(item.Properties, WIA_IPS_XRES, dpi);
            SetProperty(item.Properties, WIA_IPS_YRES, dpi);
            try { SetProperty(item.Properties, WIA_IPS_CUR_INTENT, color); } catch { }

            ImageFile img = (ImageFile)item.Transfer(wiaFormat);
            if (img != null)
            {
                images.Add(ScannerUtils.ConvertImageToBase64(img, 1));
                ScannerUtils.ReleaseCom(img);
            }

            return images;
        }

        // --------------------------------------------------------------------
        // FEEDER
        // --------------------------------------------------------------------
        private List<ScannedImage> ScanFromFeeder(Device device, string wiaFormat, int dpi, int color)
        {
            var images = new List<ScannedImage>();

            // Put device into feeder mode
            SetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT, FEEDER);

            bool hasMore = true;
            int page = 0;

            while (hasMore)
            {
                var item = device.Items[1] as Item;
                if (item == null) break;

                SetProperty(item.Properties, WIA_IPS_XRES, dpi);
                SetProperty(item.Properties, WIA_IPS_YRES, dpi);
                SetProperty(item.Properties, WIA_IPS_CUR_INTENT, color);

                try
                {
                    ImageFile img = (ImageFile)item.Transfer(wiaFormat);
                    if (img != null)
                    {
                        page++;
                        images.Add(ScannerUtils.ConvertImageToBase64(img, page));
                        ScannerUtils.ReleaseCom(img);
                    }
                }
                catch (COMException ex)
                {
                    if ((uint)ex.ErrorCode == 0x80210003)
                        break; // No more pages
                    else
                        throw;
                }

                // ---------------------------
                // FIXED FEEDER PAGE CHECK
                // ---------------------------
                int select = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_SELECT);
                int status = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_STATUS);

                hasMore = (select & FEEDER) != 0 && (status & FEED_READY) != 0;
            }

            return images;
        }

        // page → yields → scans next → yields
        private IEnumerable<ScannedImage> StreamPages(Device device, string wiaFormat, int dpi, int color, bool feederSelected)
        {
            int page = 0;

            while (true)
            {
                var item = device.Items[1] as Item;
                if (item == null) yield break;

                // resolution
                SetProperty(item.Properties, WIA_IPS_XRES, dpi);
                SetProperty(item.Properties, WIA_IPS_YRES, dpi);

                try { SetProperty(item.Properties, WIA_IPS_CUR_INTENT, color); } catch { }

                // perform transfer using a helper so any COMException is handled outside the iterator's try/catch
                bool noMorePages;

                var sw = Stopwatch.StartNew();
                ImageFile img = TryTransfer(item, wiaFormat, out noMorePages);
                sw.Stop();

                if (noMorePages) yield break;

                if (img == null) yield break;

                // log duration (replace Console.WriteLine with a logger if desired)
                page++;
                Console.WriteLine($"[Scanner] Device={(device?.DeviceID ?? "unknown")} Page={page} TransferSeconds={sw.Elapsed.TotalSeconds:N2}");


                yield return ScannerUtils.ConvertImageToBase64(img, page);

                 ScannerUtils.ReleaseCom(img);

                // If device is flatbed (feeder not selected) there is only one image — stop.
                if (!feederSelected) yield break;

                // Check feeder more pages
                int status = GetProperty(device.Properties, WIA_DPS_DOCUMENT_HANDLING_STATUS);
                if ((status & FEED_READY) == 0)
                    yield break;
            }
        }



        private void SetProperty(Properties props, int id, object value)
        {
            foreach (Property p in props)
            {
                if (p.PropertyID == id)
                {
                    p.set_Value(value);
                    break;
                }
            }
        }

        private int GetProperty(Properties props, int id)
        {
            foreach (Property p in props)
            {
                if (p.PropertyID == id)
                {
                    try { return Convert.ToInt32(p.get_Value()); }
                    catch { return 0; }
                }
            }
            return 0;
        }

        private void WriteStreamText(Stream stream, string text)
        {
            var bytes = Encoding.UTF8.GetBytes(text);
            stream.Write(bytes, 0, bytes.Length);
            stream.Flush();
        }

        private ImageFile TryTransfer(Item item, string wiaFormat, out bool noMorePages)
        {
            noMorePages = false;
            try
            {
                return (ImageFile)item.Transfer(wiaFormat);
            }
            catch (COMException ex)
            {
                // No more pages in feeder
                if ((uint)ex.ErrorCode == 0x80210003)
                {
                    noMorePages = true;
                    return null;
                }

                // rethrow other COM exceptions
                throw;
            }
        }

    }
}
