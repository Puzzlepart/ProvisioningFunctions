using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Reflection;
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
    public static class SetGraphMetadata
    {
        public static string GetLabel(MetadataField metadataField)
        {
            MemberInfo memberInfo = typeof(MetadataField).GetMember(metadataField.ToString())
                .FirstOrDefault();

            if (memberInfo == null) return null;

            DisplayAttribute attribute = (DisplayAttribute)
                memberInfo.GetCustomAttributes(typeof(DisplayAttribute), false)
                    .Single();
            return attribute.Name;
        }

        [FunctionName("SetGraphMetadata")]
        [ResponseType(typeof(SetGraphMetadataResponse))]
        [Display(Name = "Set Office 365 Group metadata", Description = "Store metadata for the Office 365 Group in the Microsoft Graph")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGraphMetadataRequest request, TraceWriter log)
        {
            try
            {
                const string extensionName = "techmikael_GenericSchema";

                string schemaKey = PropertyMapper[request.Key];
                string schemaLabel = schemaKey.Replace("Key", "Label");
                string schemaValue = schemaKey.Replace("Key", "Value");
                string label = GetLabel(request.Key);

                dynamic property = new ExpandoObject();
                ((IDictionary<string, object>)property).Add(schemaKey, request.Key.ToString());
                ((IDictionary<string, object>)property).Add(schemaLabel, label);
                ((IDictionary<string, object>)property).Add(schemaValue, request.Value);
                
                dynamic template = new ExpandoObject();
                ((IDictionary<string, object>)template).Add(extensionName, property);

                var content = new StringContent(JsonConvert.SerializeObject(template), Encoding.UTF8, "application/json");
                Uri uri = new Uri($"https://graph.microsoft.com/v1.0/groups/{request.GroupId}");
                string bearerToken = await ConnectADAL.GetBearerToken();
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PatchAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<SetGraphMetadataResponse>(new SetGraphMetadataResponse { Added = true }, new JsonMediaTypeFormatter())
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

        private static readonly Dictionary<MetadataField, string> PropertyMapper =
            new Dictionary<MetadataField, string>() { { MetadataField.GroupType, "KeyString00" },
                { MetadataField.Responsible, "KeyString01" },
                { MetadataField.StartDate, "KeyDateTime00" },
                { MetadataField.EndDate, "KeyDateTime01" } };

        public enum MetadataField
        {
            [Display(Name = "Type rom")]
            GroupType = 0,
            [Display(Name = "Kontaktperson")]
            Responsible = 1,
            [Display(Name = "Oppstartsdato")]
            StartDate = 2,
            [Display(Name = "Forventet sluttdato")]
            EndDate = 3
        }

        public class SetGraphMetadataRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Required]
            [Display(Description = "Metadata name")]
            public MetadataField Key { get; set; }

            [Required]
            [Display(Description = "Metadata value")]
            public string Value { get; set; }
        }

        public class SetGraphMetadataResponse
        {
            [Display(Description = "true/false if added")]
            public bool Added { get; set; }
        }
    }
}
