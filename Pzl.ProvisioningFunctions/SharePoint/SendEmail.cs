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
using OfficeDevPnP.Core.Utilities;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class SendEmail
    {
        [FunctionName("SendEmail")]
        [ResponseType(typeof(SendEmailResponse))]
        [Display(Name = "Send e-mail via SharePoint", Description = "Send e-mail to a user in the context of a SharePoint site")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SendEmailRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);

                MailUtility.SendEmail(clientContext, request.Recipient.Split(';'), null, request.Subject, request.Content);

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SendEmailResponse>(new SendEmailResponse { Sent = true }, new JsonMediaTypeFormatter())
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

        public class SendEmailRequest
        {
            [Required]
            [Display(Description = "URL of site to send from")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "E-mail of recipients. Separate multiple with ;")]
            public string Recipient { get; set; }

            [Required]
            [Display(Description = "Subject")]
            public string Subject { get; set; }

            [Required]
            [Display(Description = "Content, can include HTML markup")]
            public string Content { get; set; }

        }

        public class SendEmailResponse
        {
            [Display(Description = "True/false if mail was sent")]
            public bool Sent { get; set; }
        }
    }
}
