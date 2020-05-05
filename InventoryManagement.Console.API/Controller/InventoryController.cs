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
using System.Net.Mail;

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
                string loggedInUser = "", computerName = "", operatingSystem = "", ipAddress = "";

                foreach (var item in assetList)
                {
                    if (!(string.IsNullOrEmpty(item.LoggedInUser) && string.IsNullOrEmpty(item.ComputerName)
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
            catch (Exception ex)
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
                if (user != null)
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
                            if (sheet.Cells[i, 1].Value.ToString() == userRequest.ID)
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


        [HttpPost]
        [Route("sendmail")]
        public HttpResponseMessage SendMail(EmailRequest emailRequest)
        {

            try
            {

                var fromAddress = new MailAddress("prashant.parulekar1@gmail.com", "Prashant Parulekar");
                var toAddress = new MailAddress("prashant.parulekar1@gmail.com", "Prashant Parulekar");
                const string fromPassword = "";
                const string subject = "Subject";
                const string body = "Body";

                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                };

                var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = body,
                };
                var bytes = Convert.FromBase64String(emailRequest.FileArray);
                var attachment = new System.Net.Mail.Attachment(new MemoryStream(bytes), emailRequest.FileName);
                message.Attachments.Add(attachment);
                smtp.Send(message);

                return Request.CreateResponse(HttpStatusCode.OK, true);

            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }


        [HttpGet]
        [Route("getendpoints")]
        public HttpResponseMessage GetEndPoint()
        {
            try
            {
                DataTable dtEndpoint = Common.Extension.GetDataTableFromExcel("EndPoint");
                List<EndPointDashBoard> endpointList = dtEndpoint.DataTableToList<EndPointDashBoard>();


                DataTable dt = Common.Extension.GetDataTableFromExcel("Asset");
                List<Asset> assetList = dt.DataTableToList<Asset>();
                var totalCount = assetList
                        .Select(s => s.ComputerName).Distinct().Count();
                List<EndPointChart> chartDetails = new List<EndPointChart>();
                foreach (var item in endpointList)
                {
                    var keywordArray = item.Keywords.Split(',');

                    var endpointCount = assetList.Where(d =>
                    keywordArray.Any(k => d.SoftwareInstalled.Contains(k)))
                         .Select(s => s.ComputerName).Distinct().Count();

                    chartDetails.Add(new EndPointChart()
                    {
                        ID = item.ID,
                        Title = item.Title,
                        Keywords = item.Keywords,
                        AvailableCountPer = (100 / totalCount) * endpointCount,
                        AvailableCount = endpointCount,
                        TotalCount = totalCount
                    });
                }


                return Request.CreateResponse(HttpStatusCode.OK, chartDetails);
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }

        [HttpPost]
        [Route("postendpoint")]
        public HttpResponseMessage PostEndPoint(EndPointDashBoard endpointRequest)
        {

            try
            {
                using (var package = new ExcelPackage(new FileInfo(@"DB\InventoryDB.xlsx")))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets["EndPoint"];
                    if (string.IsNullOrEmpty(endpointRequest.ID))
                    {
                        var rg = sheet.Tables.First().AddRow();

                        int row = sheet.Dimension.End.Row + 1;
                        sheet.Cells[row, 1].Value = Guid.NewGuid();
                        sheet.Cells[row, 2].Value = endpointRequest.Title;
                        sheet.Cells[row, 3].Value = endpointRequest.Keywords;
                    }
                    else
                    {

                        for (int i = 0; i < sheet.Dimension.End.Row; i++)
                        {
                            if (sheet.Cells[i, 1].Value.ToString() == endpointRequest.ID)
                            {
                                sheet.Cells[i, 2].Value = endpointRequest.Title;
                                sheet.Cells[i, 3].Value = endpointRequest.Keywords;
                            }
                        }

                    }


                    package.Save();
                    return Request.CreateResponse(HttpStatusCode.OK, endpointRequest);
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.InnerException.Message));
            }

        }


        [HttpDelete]
        [Route("deleteendpoint")]
        public HttpResponseMessage DeleteEndPoint(string id)
        {

            try
            {
                using (var package = new ExcelPackage(new FileInfo(@"DB\InventoryDB.xlsx")))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets["EndPoint"];


                    for (int i = 1; i <= sheet.Dimension.End.Row; i++)
                    {
                        if (sheet.Cells[i, 1].Value.ToString() == id)
                        {
                            sheet.DeleteRow(i, 1, true);
                        }
                    }

                    package.Save();
                    return Request.CreateResponse(HttpStatusCode.OK, true);
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, new HttpError(ex.Message));
            }

        }
    }
}
