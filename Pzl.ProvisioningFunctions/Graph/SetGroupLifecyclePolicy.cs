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
using Newtonsoft.Json.Linq;
using Pzl.ProvisioningFunctions.Helpers;

namespace Pzl.ProvisioningFunctions.Graph
{
    //TODO: Change to proper GraphClient support once classification moves from Beta endpoint
    public static class SetLifecyclePolicy
    {
        [FunctionName("SetLifecyclePolicy")]
        [ResponseType(typeof(SetLifecyclePolicyResponse))]
        [Display(Name = "Set lifecycle policy to an Office 365 Group", Description = "Apply an expiration lifecyle policy to the Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetLifecyclePolicyRequest request, TraceWriter log)
        {
            try
            {
                string bearerToken = await ConnectADAL.GetBearerToken(true);
                dynamic template = new { groupId = request.GroupId };
                var content = new StringContent(JsonConvert.SerializeObject(template), Encoding.UTF8, "application/json");
                string operation = request.Operation == Operation.Add ? "addGroup" : "removeGroup";

                // Get all policies - should return 1
                Uri policyUri = new Uri("https://graph.microsoft.com/beta/groupLifecyclePolicies");
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                string policyResponse = await client.GetStringAsync(policyUri);
                JObject policies = JsonConvert.DeserializeObject<JObject>(policyResponse);
                var policyId = policies["value"][0]["id"] + "";

                Uri uri = new Uri($"https://graph.microsoft.com/beta/groupLifecyclePolicies/{policyId}/{operation}");
                var response = await client.PostAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<SetLifecyclePolicyResponse>(new SetLifecyclePolicyResponse { IsApplied = request.Operation == Operation.Add }, new JsonMediaTypeFormatter())
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

        public enum Operation
        {
            [Display(Name = "Add")]
            Add = 0,
            [Display(Name = "Remove")]
            Remove = 1
        }

        public class SetLifecyclePolicyRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Required]
            [Display(Description = "Add or Remove the group from an expiration policy")]
            public Operation Operation { get; set; }
        }

        public class SetLifecyclePolicyResponse
        {
            [Display(Description = "true/false if applied")]
            public bool IsApplied { get; set; }
        }
    }
}
