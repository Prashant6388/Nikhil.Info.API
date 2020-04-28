using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagement.Console.API.Model
{
    public class Asset
    {
        public string LoggedInUser { get; set; }
        public string ComputerName { get; set; }
        public string OperatingSystem { get; set; }
        public string IPAddress { get; set; }
        public string SoftwareInstalled { get; set; }
        public string SoftwareVersion { get; set; }
        public string PatchesInstalled { get; set; }
        public string UserAccountDetails { get; set; }
    }
}
