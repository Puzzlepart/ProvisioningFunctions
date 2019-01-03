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
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class SetSiteSharing
    {
        [FunctionName("SetSiteSharing")]
        [ResponseType(typeof(SetSiteSharingResponse))]
        [Display(Name = "Set SharePoint Sharing Level", Description = "Set who a site can be shared with")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetSiteSharingRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;
            Uri uri = new Uri(siteUrl);
            string adminUrl = uri.Scheme + "://" + uri.Authority;
            if (!adminUrl.Contains("-admin.sharepoint.com"))
            {
                adminUrl = adminUrl.Replace(".sharepoint.com", "-admin.sharepoint.com");
            }

            try
            {
                var clientContext = await ConnectADAL.GetClientContext(adminUrl, log);

                Tenant tenant = new Tenant(clientContext);
                var siteProperties = tenant.GetSitePropertiesByUrl(siteUrl, false);
                clientContext.Load(siteProperties, s => s.SharingCapability);
                siteProperties.Context.ExecuteQueryRetry();

                bool sharingUpdated = false;
                SharingCapabilities externalShareCapabilities;
                switch (request.SharingCapabilities)
                {
                    case ExternalSharingCapabilities.Disabled:
                        externalShareCapabilities = SharingCapabilities.Disabled;
                        break;
                    case ExternalSharingCapabilities.ExistingExternalUserSharingOnly:
                        externalShareCapabilities = SharingCapabilities.ExistingExternalUserSharingOnly;
                        break;
                    case ExternalSharingCapabilities.ExternalUserAndGuestSharing:
                        externalShareCapabilities = SharingCapabilities.ExternalUserAndGuestSharing;
                        break;
                    case ExternalSharingCapabilities.ExternalUserSharingOnly:
                        externalShareCapabilities = SharingCapabilities.ExternalUserSharingOnly;
                        break;
                    default:
                        externalShareCapabilities = SharingCapabilities.Disabled;
                        break;
                }
                if (siteProperties.SharingCapability != externalShareCapabilities)
                {
                    sharingUpdated = true;
                    siteProperties.SharingCapability = externalShareCapabilities;
                    siteProperties.Update();
                    clientContext.ExecuteQueryRetry();
                }

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetSiteSharingResponse>(new SetSiteSharingResponse{ UpdatedSharing = sharingUpdated }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public enum ExternalSharingCapabilities
        {
            [Display(Name = "Don't allow sharing outside your organization")]
            Disabled,
            [Display(Name = "Allow users to invite and share with authenticated external users")]
            ExternalUserSharingOnly,
            [Display(Name = "Allow sharing to authenticated external users and using anonymous access links")]
            ExternalUserAndGuestSharing,
            [Display(Name = "Allow sharing only with the external users that already exist in your organization's directory")]
            ExistingExternalUserSharingOnly
        }

        public class SetSiteSharingRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Sharing outside your organization")]
            public ExternalSharingCapabilities SharingCapabilities { get; set; }
        }

        public class SetSiteSharingResponse
        {
            [Display(Description = "True/false if sharing capabilities was updated")]
            public bool UpdatedSharing { get; set; }
        }
    }
}
