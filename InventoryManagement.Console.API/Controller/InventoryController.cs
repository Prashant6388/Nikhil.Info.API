using InventoryManagement.Console.API.Common;
using InventoryManagement.Console.API.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using OfficeOpenXml;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using System.IO;

namespace InventoryManagement.Console.API.Controller
{
    [RoutePrefix("api/inventory")]
    public class InventoryController : ApiController
    {
       
        [HttpGet]
        [Route("getassets")]
        public HttpResponseMessage GetAllAssets()
        {         
             try
            {
                DataTable dt = Common.Extension.GetDataTableFromExcel("Asset");
                List<Asset> assetList = dt.DataTableToList<Asset>();  
                string loggedInUser = "", computerName="", operatingSystem="", ipAddress = "";

                foreach (var item in assetList)
                {
                    if(!(string.IsNullOrEmpty(item.LoggedInUser) && string.IsNullOrEmpty(item.ComputerName) 
                        && string.IsNullOrEmpty(item.OperatingSystem) && string.IsNullOrEmpty(item.IPAddress)))
                    {
                        loggedInUser = item.LoggedInUser;
                        computerName = item.ComputerName;
                        operatingSystem = item.OperatingSystem;
                        ipAddress = item.IPAddress;
                    }
                    else
                    {
                        item.LoggedInUser = loggedInUser;
                        item.ComputerName = computerName;
                        item.OperatingSystem = operatingSystem;
                        item.IPAddress = ipAddress;
                    }
                }
                return Request.CreateResponse(HttpStatusCode.OK, assetList);
            }
            catch(Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }
           
        }

        [HttpGet]
        [Route("getvaservers")]
        public HttpResponseMessage GetVAServers()
        {

            try
            {
                DataTable dt = Common.Extension.GetDataTableFromExcel("VAServer");
                List<VAServer> vaserverList = dt.DataTableToList<VAServer>();
                return Request.CreateResponse(HttpStatusCode.OK, vaserverList);
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }

        [HttpGet]
        [Route("getptnetworkservices")]
        public HttpResponseMessage GetPTNetworkServices()
        {

            try
            {
                DataTable dt = Common.Extension.GetDataTableFromExcel("PTNetworkService");
                List<PTNetworkService> ptList = dt.DataTableToList<PTNetworkService>();
                return Request.CreateResponse(HttpStatusCode.OK, ptList);
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }

        [HttpGet]
        [Route("getapplicationsecurity")]
        public HttpResponseMessage GetApplicationSecurity()
        {

            try
            {
                DataTable dt = Common.Extension.GetDataTableFromExcel("ApplicationSecurity");
                List<ApplicationSecurity> ptList = dt.DataTableToList<ApplicationSecurity>();
                return Request.CreateResponse(HttpStatusCode.OK, ptList);
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }

        [HttpGet]
        [Route("getusers")]
        public HttpResponseMessage GetUsers()
        {

            try
            {
                DataTable dt = Common.Extension.GetDataTableFromExcel("User");
                List<User> ptList = dt.DataTableToList<User>();
                return Request.CreateResponse(HttpStatusCode.OK, ptList);
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }


        [HttpPost]
        [Route("login")]
        public HttpResponseMessage Login(LoginRequest loginRequest)
        {

            try
            {
                DataTable dt = Common.Extension.GetDataTableFromExcel("User");
                List<User> ptList = dt.DataTableToList<User>();
                var user = ptList.Where(d => d.Email == loginRequest.Email && d.Password == loginRequest.Password).FirstOrDefault();
                if (user!=null)
                {
                    return Request.CreateResponse(HttpStatusCode.OK, user);
                }
                return Request.CreateResponse(HttpStatusCode.Forbidden, "Login Failed");
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }

        [HttpPost]
        [Route("postuser")]
        public HttpResponseMessage PostUser(UserRequest userRequest)
        {

            try
            {
                DataTable dt = Common.Extension.GetDataTableFromExcel("User");
                List<User> ptList = dt.DataTableToList<User>();
                using (var package = new ExcelPackage(new FileInfo(@"DB\InventoryDB.xlsx")))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets["User"];
                    if (string.IsNullOrEmpty(userRequest.ID))
                    {
                        var rg = sheet.Tables.First().AddRow();

                        int row = sheet.Dimension.End.Row + 1;
                        sheet.Cells[row, 1].Value = Guid.NewGuid();
                        sheet.Cells[row, 2].Value = userRequest.Name;
                        sheet.Cells[row, 3].Value = userRequest.Email;
                        sheet.Cells[row, 4].Value = userRequest.Password;
                        sheet.Cells[row, 4].Value = userRequest.ContactNo;
                    }
                    else
                    {
                        
                        for (int i = 0; i < sheet.Dimension.End.Row; i++)
                        {
                            if(sheet.Cells[i, 1].Value.ToString() == userRequest.ID)
                            {
                                sheet.Cells[i, 2].Value = userRequest.Name;
                                sheet.Cells[i, 3].Value = userRequest.Email;
                                sheet.Cells[i, 4].Value = userRequest.Password;
                                sheet.Cells[i, 4].Value = userRequest.ContactNo;
                            }
                        }
                        
                    }
                   
                    
                    package.Save();
                    return Request.CreateResponse(HttpStatusCode.OK, userRequest);
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }

    }
}
