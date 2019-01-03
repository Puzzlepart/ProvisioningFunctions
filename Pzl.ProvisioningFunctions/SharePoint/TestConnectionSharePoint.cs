using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class TestConnectionSharePoint
    {
        [FunctionName("TestConnectionSharePointSharePoint")]
        [ResponseType(typeof(TestConnectionSharePointResponse))]
        [Display(Name = "TestConnectionSharePointSharePoint", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]TestConnectionSharePointRequest request, TraceWriter log)
        {           
            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }

                var clientContext = await ConnectADAL.GetClientContext(request.SiteURL, log);
                var web = clientContext.Web;
                clientContext.Load(web);
                clientContext.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<TestConnectionSharePointResponse>(new TestConnectionSharePointResponse { WebTitle = web.Title }, new JsonMediaTypeFormatter())
                });
            }
            catch(ArgumentException ae)
            {
                log.Error($"Error:  {ae.Message }\n\n{ae.StackTrace}");
                var response = new ExpandoObject();
                ((IDictionary<string, object>)response).Add("message", ae.Message);
                ((IDictionary<string, object>)response).Add("statusCode", HttpStatusCode.BadRequest);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new ObjectContent<ExpandoObject>(response, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                var response = new ExpandoObject();
                ((IDictionary<string, object>)response).Add("message", e.Message);
                ((IDictionary<string, object>)response).Add("statusCode", HttpStatusCode.ServiceUnavailable);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<ExpandoObject>(response, new JsonMediaTypeFormatter())
                });
            }
        }

        public class TestConnectionSharePointRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
        }

        public class TestConnectionSharePointResponse
        {
            [Display(Description = "")]
            public string WebTitle { get; set; }
        }
    }
}
