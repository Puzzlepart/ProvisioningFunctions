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
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGroupNamePrefixRequest request, TraceWriter log)
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
                clientContext.Load(associatedOwnerGroup, group => group.Title);
                clientContext.Load(associatedMemberGroup, group => group.Title);
                clientContext.Load(associatedVisitorGroup, group => group.Title);
                clientContext.ExecuteQuery();

                var associatedOwnerGroupTitle = $"({request.Prefix}) {associatedOwnerGroup.Title}";
                if (!associatedOwnerGroup.Title.StartsWith(request.Prefix))
                {
                    log.Info($"Setting title of AssociatedOwnerGroup to {associatedOwnerGroupTitle}.");
                    web.AssociatedOwnerGroup.Title = associatedOwnerGroupTitle;
                    web.AssociatedOwnerGroup.Update();
                }

                var associatedMemberGroupTitle = $"({request.Prefix}) {associatedMemberGroup.Title}";
                if (!associatedMemberGroup.Title.StartsWith(request.Prefix))
                {
                    log.Info($"Setting title of AssociatedOwnerGroup to {associatedMemberGroupTitle}.");
                    web.AssociatedMemberGroup.Title = associatedMemberGroupTitle;
                    web.AssociatedMemberGroup.Update();
                }

                var associatedVisitorGroupTitle = $"({request.Prefix}) {associatedVisitorGroup.Title}";
                if (!associatedVisitorGroup.Title.StartsWith(request.Prefix))
                {
                    log.Info($"Setting title of AssociatedOwnerGroup to {associatedVisitorGroupTitle}.");
                    web.AssociatedVisitorGroup.Title = associatedVisitorGroupTitle;
                    web.AssociatedVisitorGroup.Update();
                }

                if (clientContext.HasPendingRequest)
                {
                    clientContext.ExecuteQuery();
                }

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetGroupNamePrefixResponse>(new SetGroupNamePrefixResponse {
                        AssociatedOwnerGroupTitle = associatedOwnerGroupTitle,
                        AssociatedMemberGroupTitle = associatedMemberGroupTitle,
                        AssociatedVisitorGroupTitle = associatedVisitorGroupTitle,
                    }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.InternalServerError)
                {
                    Content = new ObjectContent<SetGroupNamePrefixResponse>(new SetGroupNamePrefixResponse { Message = e.Message }, new JsonMediaTypeFormatter())
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

        public class SetGroupNamePrefixResponse {
            public string Message { get; set; }
            public string AssociatedOwnerGroupTitle { get; set; }
            public string AssociatedMemberGroupTitle { get; set; }
            public string AssociatedVisitorGroupTitle { get; set; }
        }
    }
}
