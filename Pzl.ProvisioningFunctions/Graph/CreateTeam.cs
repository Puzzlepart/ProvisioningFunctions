using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;
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
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.Graph
{
    public static class CreateTeam
    {
        [FunctionName("CreateTeam")]
        [ResponseType(typeof(CreateTeamResponse))]
        [Display(Name = "Create a Team for a Office 365 Group", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]CreateTeamRequest request, TraceWriter log)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.GroupId))
                {
                    throw new ArgumentException("Parameter cannot be null", "GroupId");
                }

                dynamic team = new ExpandoObject();
                var content = new StringContent(JsonConvert.SerializeObject(team), Encoding.UTF8, "application/json");
                log.Info(JsonConvert.SerializeObject(team));
                Uri uri = new Uri($"https://graph.microsoft.com/beta/groups/{request.GroupId}/team");
                log.Info(uri.AbsoluteUri);
                string bearerToken = await ConnectADAL.GetBearerToken();
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PutAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    dynamic responseJson = JObject.Parse(responseBody);
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<CreateTeamResponse>(new CreateTeamResponse { Created = true, TeamUrl = responseJson.webUrl }, new JsonMediaTypeFormatter())
                    });
                }
                string responseMsg = await response.Content.ReadAsStringAsync();
                log.Info(responseMsg);
                dynamic errorMsg = JsonConvert.DeserializeObject(responseMsg);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<CreateTeamResponse>(new CreateTeamResponse { Created = false, ErrorMessage = errorMsg }, new JsonMediaTypeFormatter())
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

        public class PutTeamResponse
        {
            public string id { get; set; }
            public string webId { get; set; }
        }


        public class CreateTeamRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
        }

        public class CreateTeamResponse
        {
            [Display(Description = "True/false if created")]
            public bool Created { get; set; }
            [Display(Description = "Team URL")]
            public string TeamUrl { get; set; }
            [Display(Description = "Error message if applicable")]
            public string ErrorMessage { get; set; }
        }
    }
}
