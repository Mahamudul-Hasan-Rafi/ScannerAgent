using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerAgent.Model
{
    internal class ScanResponse
    {
        public int statusCode { get; set; }
        
        public List<ScannedImage> images { get; set; }

        public String status { get; set; }
    }
}
