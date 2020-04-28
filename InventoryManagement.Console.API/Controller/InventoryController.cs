using InventoryManagement.Console.API.Common;
using InventoryManagement.Console.API.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

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
    }
}
