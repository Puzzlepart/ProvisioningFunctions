using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.SharePoint
{
    public static class SetGroupNamePrefix
    {
        [FunctionName("SetGroupNamePrefix")]
        [ResponseType(typeof(SetGroupNamePrefixResponse))]
        [Display(Name = "Set group name prefix", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get")]SetGroupNamePrefixRequest request, TraceWriter log)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.Prefix))
                {
                    throw new ArgumentException("Parameter cannot be null", "Prefix");
                }

                var clientContext = await ConnectADAL.GetClientContext(request.SiteURL, log);

                var web = clientContext.Web;
                var associatedOwnerGroup = web.AssociatedOwnerGroup;
                var associatedMemberGroup = web.AssociatedMemberGroup;
                var associatedVisitorGroup = web.AssociatedVisitorGroup;
                web.Context.Load(associatedOwnerGroup);
                web.Context.Load(associatedMemberGroup);
                web.Context.Load(associatedVisitorGroup);
                web.Context.ExecuteQueryRetry();

                associatedOwnerGroup.Title = $"{request.Prefix}: ${associatedOwnerGroup.Title}";
                associatedMemberGroup.Title = $"{request.Prefix}: ${associatedMemberGroup.Title}";
                associatedVisitorGroup.Title = $"{request.Prefix}: ${associatedVisitorGroup.Title}";

                associatedOwnerGroup.Update();
                associatedMemberGroup.Update();
                associatedVisitorGroup.Update();

                web.Context.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetGroupNamePrefixResponse>(new SetGroupNamePrefixResponse { }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetGroupNamePrefixResponse>(new SetGroupNamePrefixResponse { }, new JsonMediaTypeFormatter())
                });
            }
        }

        public class SetGroupNamePrefixRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
            [Required]
            [Display(Description = "Permission level for Owners")]
            public string Prefix { get; set; }
        }

        public class SetGroupNamePrefixResponse { }
    }
}
