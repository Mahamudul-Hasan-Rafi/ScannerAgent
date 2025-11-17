
using ScannerAgent.Model;
using ScannerAgent.Services;
using System.Collections.Generic;
using System.Threading.Tasks;
using WIA;

namespace ScannerAgent.Controllers
{
    public class ScannerController
    {
        private readonly ScannerService _scannerService = new ScannerService();

        // Async scan method
        public async Task<(List<ScannedImage> images, string status, int statusCode)> ScanAsync(ScanRequest request)
        {
            var result = await _scannerService.ScanAsync(request);
            return result;
        }



        public List<DeviceInformation> Devices()
        {
            return _scannerService.ListDevices();
        }
    }
}
