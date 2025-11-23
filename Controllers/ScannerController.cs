
using ScannerAgent.Model;
using ScannerAgent.Services;
using System.Collections.Generic;
using System.Net;
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

        public async Task StreamScannedDocumentAsync(HttpListenerContext context, ScanRequest request)
        {
            await _scannerService.StreamScanAsync(context, request);
        }

        public DeviceListResponse Devices()
        {
            return _scannerService.ListDevices();
        }
    }
}
