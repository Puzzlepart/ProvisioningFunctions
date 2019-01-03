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
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class AddUserToGroup
    {
        [FunctionName("AddUserToGroup")]
        [ResponseType(typeof(AddUserToGroupResponse))]
        [Display(Name = "Add user to SharePoint group", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]AddUserToGroupRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.Title))
                {
                    throw new ArgumentException("Parameter cannot be null", "Title");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);


                var web = clientContext.Web;
                var user = web.EnsureUser(request.Title);
                Microsoft.SharePoint.Client.Group siteGroup = null;

                switch(request.AssociatedGroup)
                {
                    case AssociatedGroup.Member:
                        {
                            siteGroup = web.AssociatedMemberGroup;
                        }
                        break;
                    case AssociatedGroup.Owner:
                        {
                            siteGroup = web.AssociatedOwnerGroup;
                        }
                        break;
                    case AssociatedGroup.Visitor:
                        {
                            siteGroup = web.AssociatedVisitorGroup;
                        }
                        break;
                }


                if (siteGroup != null)
                {
                    web.Context.Load(siteGroup);
                    web.Context.Load(user);
                    web.Context.ExecuteQueryRetry();

                    siteGroup.Users.AddUser(user);
                    web.Context.ExecuteQueryRetry();

                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<AddUserToGroupResponse>(new AddUserToGroupResponse { UserAdded = true}, new JsonMediaTypeFormatter())
                    });
                } else
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<AddUserToGroupResponse>(new AddUserToGroupResponse { UserAdded = false }, new JsonMediaTypeFormatter())
                    });
                }
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<AddUserToGroupResponse>(new AddUserToGroupResponse { UserAdded = false }, new JsonMediaTypeFormatter())
                });
            }
        }

        public enum AssociatedGroup
        {
            Member,
            Owner,
            Visitor
        }

        public class AddUserToGroupRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
            [Required]
            [Display(Description = "User")]
            public string Title { get; set; }
            [Required]
            [Display(Description = "Associated group")]
            public AssociatedGroup AssociatedGroup { get; set; }

        }

        public class AddUserToGroupResponse {

            [Display(Description = "Was user added to group")]
            public bool UserAdded { get; set; }
        }
    }
}
