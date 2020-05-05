using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagement.Console.API.Model
{
    public class User
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string ContactNo{ get; set; }
    }

    public class LoginRequest
    {       
        public string Email { get; set; }        
        public string Password { get; set; }       
    }

    public class UserRequest
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string ContactNo { get; set; }
    }
    public class EmailRequest
    {
        public string FileArray { get; set; }
        public string FileName { get; set; }
    }
}
