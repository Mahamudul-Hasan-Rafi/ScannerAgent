using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerAgent.Model
{
    public class ScannedImage
    {
        public int PageNumber { get; set; }
        public string Base64Data { get; set; }
        public long Size { get; set; }
        public string Format { get; set; }
    }
}
