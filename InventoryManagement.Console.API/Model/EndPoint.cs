using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagement.Console.API.Model
{
    public class EndPointDashBoard
    {
        public string ID { get; set; }
        public string Title { get; set; }
        public string Keywords { get; set; }
    }

    public class EndPointChart
    {
        public string ID { get; set; }
        public string Title { get; set; }
        public string Keywords { get; set; }
        public int TotalCount { get; set; }
        public int AvailableCount { get; set; }
        public int AvailableCountPer { get; set; }
    }
}
