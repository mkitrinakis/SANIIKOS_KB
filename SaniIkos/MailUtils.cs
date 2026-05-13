using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;

using System.Threading.Tasks;
using System.Net;
using PnP.Framework;
using System.Security.Cryptography;
using AuthenticationManager = PnP.Framework.AuthenticationManager;
using PnP.Core.Model.SharePoint;
using static System.Net.WebRequestMethods;
using FieldUserValue = Microsoft.SharePoint.Client.FieldUserValue;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Identity.Client;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Azure.Identity;
using PnP.Framework.Provisioning.Model;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Runtime.ConstrainedExecution;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace SaniIkos
{
    internal class MailUtils
    {
        static string[] scopes = new[] { "https://graph.microsoft.com/.default" };
        const string clientID = "048730f5-e744-4f81-8055-596b705903ee";
        const string tenantID = "5f5b00f6-41ee-4bd1-8944-9cb3b1e3fef2";
        const string certificationPath = "C:\\Certs\\KBSchedule.lbriresorts.eu.pfx";
        const string certificationPassword = "lbri_123!2";

        public static void sendMail()
        {

            X509Certificate2 certificate = new X509Certificate2(certificationPath, certificationPassword);
            ClientCertificateCredential credential = new ClientCertificateCredential(
    tenantId: tenantID,
    clientId: clientID,
    clientCertificate: certificate
);

            //ConfidentialClientApplication app = (
            //ConfidentialClientApplication)ConfidentialClientApplicationBuilder.Create(clientID)
            //    .WithCertificate(certificate)
            //    .WithTenantId(tenantID)
            //    .Build();

            //var result = app.AcquireTokenForClient(scopes).ExecuteAsync().GetAwaiter().GetResult();
            //string accessToken = GetAccessToken(app);
            //callGraph(accessToken);
            callGraph2(credential); 
        }

        public static string GetAccessToken(ConfidentialClientApplication app)
        {
            return GetAccessTokenAsync(app).GetAwaiter().GetResult();
        }

        private static async Task<string> GetAccessTokenAsync(ConfidentialClientApplication app)
        {
            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            return result.AccessToken;
        }


        private static async void callGraph(string accessToken)
        {
            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);

                var mail = new
                {
                    message = new
                    {
                        subject = "Test email",
                        body = new
                        {
                            contentType = "HTML",
                            content = "<p>Hello from KB, this is just a test mail from KB. Alexandros please if you receive it forward to mkitrinakis@inttrust.gr</p>"
                        },
                        toRecipients = new[]
                        {
                new
                {
                    emailAddress = new { address = "atzimpilis@saniikos.com" }
                }
            }
                    },
                    saveToSentItems = "true"
                };

                var json = JsonConvert.SerializeObject(mail);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync(
                    "https://graph.microsoft.com/v1.0/users/noreply@company.com/sendMail",
                    content);

                response.EnsureSuccessStatusCode();
            }
        }

        private static async void callGraph2(ClientCertificateCredential credential)
        {
            GraphServiceClient graph; 
            try
            {
                graph = new GraphServiceClient(credential);
                Console.WriteLine("Graph client created");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw; // keep it so debugger stops here
            }
           // GraphServiceClient graph = new GraphServiceClient(credential);
            var body = new SendMailPostRequestBody
            {
                Message = new Message
                {
                    Subject = "Test",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = "<p>Hello from KB, this is just a test mail from KB. Alexandros please if you receive it forward to mkitrinakis@inttrust.gr</p>"
                    },
                    ToRecipients = new List<Recipient>
        {
            new Recipient
            {
                EmailAddress = new EmailAddress { Address = "atzimpilis@saniikos.com" }
            }
        }
                },
                SaveToSentItems = true
            };
            try
            {
                await graph.Users["sa_intt_sadm@lbriresorts.eu"].SendMail.PostAsync(body);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message); 
            }

        }
    }
}
