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
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.Graph
{
    //TODO: Change to proper GraphClient support once classification moves from Beta endpoint
    public static class SetGroupClassification
    {
        [FunctionName("SetGroupClassification")]
        [ResponseType(typeof(SetGroupClassificationResponse))]
        [Display(Name = "Set classification value for an Office 365 Group", Description = "Set the classification value for the Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGroupClassificationRequest request, TraceWriter log)
        {
            try
            {
                string bearerToken = await ConnectADAL.GetBearerToken();
                dynamic template = new { classification = request.Classification };
                var content = new StringContent(JsonConvert.SerializeObject(template), Encoding.UTF8, "application/json");
                Uri uri = new Uri($"https://graph.microsoft.com/beta/groups/{request.GroupId}");
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PatchAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<SetGroupClassificationResponse>(new SetGroupClassificationResponse{ IsUpdated= true}, new JsonMediaTypeFormatter())
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

        public class SetGroupClassificationRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
            [Required]
            [Display(Description = "Classification label to be used")]
            public string Classification { get; set; }
        }

        public class SetGroupClassificationResponse
        {
            [Display(Description = "true/false if set")]
            public bool IsUpdated { get; set; }
        }
    }
}
