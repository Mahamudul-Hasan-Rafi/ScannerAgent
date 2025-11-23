using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerAgent.Model
{
    public class DeviceListResponse
    {
        public int statusCode { get; set; }

        public List<DeviceInformation> devices { get; set; }

        public string status { get; set; }
    }
}
