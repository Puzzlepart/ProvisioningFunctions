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
    public static class GetSharePointGroupUsers
    {
        [FunctionName("GetSharePointGroupUsers")]
        [ResponseType(typeof(GetSharePointGroupUsersResponse))]
        [Display(Name = "Get users in a SharePoint group", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]GetSharePointGroupUsersRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.Group))
                {
                    throw new ArgumentException("Parameter cannot be null", "Title");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);

                var regex = new Regex("[^a-zA-Z0-9 -]");
                var cleanGroupName = regex.Replace(request.Group, "");
                cleanGroupName = Regex.Replace(cleanGroupName, @"\s+", " ");

                var web = clientContext.Web;
                var group = web.SiteGroups.GetByName(cleanGroupName);
                clientContext.Load(group, g => g.Users);
                web.Context.ExecuteQueryRetry();

                var emails = String.Join(";", group.Users.Where(u => !String.IsNullOrWhiteSpace(u.Email)).Select(u => u.Email).ToArray());
                

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<GetSharePointGroupUsersResponse>(new GetSharePointGroupUsersResponse { Group = cleanGroupName, Emails = emails }, new JsonMediaTypeFormatter())
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

        public class GetSharePointGroupUsersRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "")]
            public string Group { get; set; }
        }

        public class GetSharePointGroupUsersResponse { 
            public string Group { get; set; }
            public string Emails { get; set; }
        }
    }
}
