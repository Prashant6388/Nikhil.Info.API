using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagement.Console.API.Model
{
    public class ApplicationSecurity
    {
        public string Application { get; set; }
        public string Vulnerability { get; set; }
        public string Severity { get; set; }
        public string Severity1 { get; set; }
        public string Status { get; set; }
    }
}
