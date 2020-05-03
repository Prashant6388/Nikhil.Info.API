using InventoryManagement.Console.API.Common;
using InventoryManagement.Console.API.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace InventoryManagement.Console.API.Controller
{
    [RoutePrefix("api/product")]
    public class ProductController : ApiController
    {
        Product[] products = new Product[]
       {
            new Product { Id = 1, Name = "Tomato Soup", Category = "Groceries", Price = 1 },
            new Product { Id = 2, Name = "Yo-yo", Category = "Toys", Price = 3.75M },
            new Product { Id = 3, Name = "Hammer", Category = "Hardware", Price = 16.99M }
       };

        [HttpGet]
        [Route("")]        
        public IEnumerable<Product> GetAllProducts()
        {
            //var package = new ExcelPackage(new FileInfo(@"DB\InventoryDB.xlsx"));
            //ExcelWorksheet sheet = package.Workbook.Worksheets[0];

            //var table = sheet.Tables.First();


            DataTable dt = Common.Extension.GetDataTableFromExcel("User");
            List<User> userList = dt.DataTableToList<User>();

            using (var package = new ExcelPackage(new FileInfo(@"DB\InventoryDB.xlsx")))
            {
           
                ExcelWorksheet sheet = package.Workbook.Worksheets["User"];
                //var rg = sheet.Tables.First().AddRow();
               
                //rg.LoadFromText("3,Prashant PG,Prashant@gmail.com,1234");

                int row = sheet.Dimension.End.Row + 1;
                sheet.Cells[row, 1].Value = 3;
                sheet.Cells[row, 2].Value = "Prashant PGGFDGFDGFD";
                sheet.Cells[row, 3].Value = "Prashant@gmail.com";
                sheet.Cells[row, 4].Value = "1234";
                package.Save();         

            }
              
            return products;
        }


        [HttpGet]
        [Route("{id}")]
        public Product GetProductById(int id)
        {



            var product = products.FirstOrDefault((p) => p.Id == id);
            if (product == null)
            {
                throw new HttpResponseException(HttpStatusCode.NotFound);
            }
            return product;
        }

        [HttpGet]
        [Route("getproductbycategory/{category}")]
        public IEnumerable<Product> GetProductsByCategory(string category)
        {
            return products.Where(p => string.Equals(p.Category, category,
                    StringComparison.OrdinalIgnoreCase));
        }
    }
}
