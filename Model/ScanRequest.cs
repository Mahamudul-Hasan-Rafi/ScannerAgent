using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerAgent.Model
{
    public class ScanRequest
    {
        public string Mode { get; set; }      // "flatbed", "adf", or "auto"
        public string Format { get; set; }    // "jpeg", "png", "tiff", "bmp"
        public int? Dpi { get; set; }         // e.g., 300
        public int? Color { get; set; }       // 1=text, 2=grayscale, 4=color
        public string DeviceId { get; set; }  // optional, scanner device ID
    }
}
/**
 * 
 * {
  "Mode": "adf",
  "Format": "png",
  "Dpi": 300,
  "Color": 4,
  "DeviceId": "xyz"
}
*/