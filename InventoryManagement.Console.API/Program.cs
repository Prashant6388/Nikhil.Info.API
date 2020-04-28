using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.SelfHost;

namespace InventoryManagement.Console.API
{
    class Program
    {
        static void Main(string[] args)
        {
            var config = new HttpSelfHostConfiguration("http://localhost:7500");
            EnableCorsAttribute cors = new EnableCorsAttribute("http://localhost:4200", "*", "GET,POST");
            config.EnableCors(cors);

            //config.Routes.MapHttpRoute(
            //    "API Default", "api/{controller}/{id}",
            //    new { id = RouteParameter.Optional });
            config.MapHttpAttributeRoutes();
            using (HttpSelfHostServer server = new HttpSelfHostServer(config))
            {
                server.OpenAsync().Wait();
                System.Console.WriteLine("Press Enter to quit.");
                System.Console.ReadLine();
            }
        }
    }
}
