using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagement.Console.API.Model
{
    public class User
    {
        public int ID { get; set; }
        public string UserName { get; set; }
        public string UserEmail { get; set; }
        public string UserPassword { get; set; }
        public string UserContact{ get; set; }
        public bool Active { get; set; }
    }
}
