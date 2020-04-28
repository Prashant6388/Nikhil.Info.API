using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagement.Console.API.Model
{
    public class VAServer
    {
        public string Datacenter { get; set; }
        public string Hostname { get; set; }
        public string IPAddress { get; set; }
        public string Vulnerability { get; set; }
        public string Severity { get; set; }
        public string Severity1 { get; set; }
        public string Status { get; set; }

    }
}
