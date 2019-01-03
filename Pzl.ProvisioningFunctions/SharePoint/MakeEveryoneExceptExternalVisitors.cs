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
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class MakeEveryoneExceptExternalVisitors
    {
        [FunctionName("MakeEveryoneExceptExternalVisitors")]
        [ResponseType(typeof(MakeEveryoneExceptExternalVisitorsResponse))]
        [Display(Name = "Make Everyone but external users visitor", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]MakeEveryoneExceptExternalVisitorsRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                bool everyOneExceptExternalAddedToVisitors = await MakeEveryoneVisitor(log, siteUrl);

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<MakeEveryoneExceptExternalVisitorsResponse>(new MakeEveryoneExceptExternalVisitorsResponse { EveryOneExceptExternalAddedToVisitors = everyOneExceptExternalAddedToVisitors }, new JsonMediaTypeFormatter())
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

        private static async Task<bool> MakeEveryoneVisitor(TraceWriter log, string siteUrl)
        {
            var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
            const string everyoneIdent = "c:0-.f|rolemanager|spo-grid-all-users/";
            bool everyOneExceptExternalAddedToVisitors = false;

            var web = clientContext.Web;
            var visitorsGroup = web.AssociatedVisitorGroup;
            var siteUsers = web.SiteUsers;

            clientContext.Load(visitorsGroup);
            clientContext.Load(siteUsers);
            clientContext.ExecuteQueryRetry();

            foreach (User user in siteUsers)
            {
                if (user.LoginName.StartsWith(everyoneIdent))
                {
                    if(!web.IsUserInGroup(visitorsGroup.Title, user.LoginName))
                    {
                        web.AddUserToGroup(visitorsGroup, user);
                        everyOneExceptExternalAddedToVisitors = true;
                    }
                }
            }
            
            return everyOneExceptExternalAddedToVisitors;
        }

        public class MakeEveryoneExceptExternalVisitorsRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
        }

        public class MakeEveryoneExceptExternalVisitorsResponse
        {
            [Display(Description = "Everyone but external users was added to visitor group")]
            public bool EveryOneExceptExternalAddedToVisitors { get; set; }
        }
    }
}
