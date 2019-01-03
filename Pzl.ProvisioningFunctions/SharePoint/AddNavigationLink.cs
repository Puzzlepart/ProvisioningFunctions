using System;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Enums;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class AddNavigationLink
    {
        [FunctionName("AddNavigationLink")]
        [ResponseType(typeof(AddNavigationResponse))]
        [Display(Name = "Add navigation link", Description = "Adds a navigational link to the site.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]AddNavigationRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                var web = clientContext.Web;
                if (string.IsNullOrWhiteSpace(request.ParentTitle)) request.ParentTitle = string.Empty;
                bool isExternal = !request.NavigationURL.Contains("sharepoint.com");
                var node = web.AddNavigationNode(request.Title, new Uri(request.NavigationURL), request.ParentTitle, request.Type, isExternal, !request.AddFirst);

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<AddNavigationResponse>(new AddNavigationResponse { NavigationAdded = node != null }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class AddNavigationRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Navigation title")]
            public string Title { get; set; }

            [Display(Description = "Navigation URL")]
            public string NavigationURL { get; set; }

            [Display(Description = "Parent navigation title")]
            public string ParentTitle { get; set; }

            [Required]
            [Display(Description = "Type of navigation")]
            public NavigationType Type { get; set; }

            [Display(Description = "Add as first navigation node")]
            public bool AddFirst { get; set; }
        }

        public class AddNavigationResponse
        {
            [Display(Description = "True if navigation node was added")]
            public bool NavigationAdded { get; set; }
        }
    }
}
