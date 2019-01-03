using System;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.Graph
{
    //TODO: Change to proper GraphClient support once classification moves from Beta endpoint
    public static class RenewLifecyclePolicy
    {
        [FunctionName("RenewLifecyclePolicy")]
        [ResponseType(typeof(RenewLifecyclePolicyResponse))]
        [Display(Name = "Renew expiration for an Office 365 Group", Description = "Extend lifecycle expiration for the Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]RenewLifecyclePolicyRequest request, TraceWriter log)
        {
            try
            {
                string bearerToken = await ConnectADAL.GetBearerToken();
                dynamic template = new { groupId = request.GroupId};
                var content = new StringContent(JsonConvert.SerializeObject(template), Encoding.UTF8, "application/json");
                Uri uri = new Uri($"https://graph.microsoft.com/beta/groupLifecyclePolicies/renewGroup");
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PostAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<RenewLifecyclePolicyResponse>(new RenewLifecyclePolicyResponse{ IsExtended= true}, new JsonMediaTypeFormatter())
                    });
                }

                string responseMsg = await response.Content.ReadAsStringAsync();
                dynamic errorMsg = JsonConvert.DeserializeObject(responseMsg);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<object>(errorMsg, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class RenewLifecyclePolicyRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
        }

        public class RenewLifecyclePolicyResponse
        {
            [Display(Description = "true/false if set")]
            public bool IsExtended { get; set; }
        }
    }
}
