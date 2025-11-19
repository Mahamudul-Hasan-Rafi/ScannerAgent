using ScannerAgent.Model;
using ScannerAgent.Utils;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WIA;

namespace ScannerAgent.Services
{
    public class ScannerService
    {
        private readonly SemaphoreSlim _scanSemaphore = new SemaphoreSlim(1, 1);
        private readonly TimeSpan _scanTimeout = TimeSpan.FromMinutes(2);

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

                return tcs.Task.Result;
            }
            finally
            {
                _scanSemaphore.Release();
            }
        }

        public List<DeviceInformation> ListDevices()
        {
            var devicesList = new List<DeviceInformation>();

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

                    devicesList.Add(new DeviceInformation
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
            }

            return devicesList;
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
    }
}
