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
//using Microsoft.Graph.Models;
//using Microsoft.Graph.Users.Item.SendMail;
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
using PnP.Framework.Utilities;

namespace Alerts
{
    internal class MailUtils
    {
        string[] scopes = new[] { "https://graph.microsoft.com/.default" };
        //const string clientID = "048730f5-e744-4f81-8055-596b705903ee";
        //const string tenantID = "5f5b00f6-41ee-4bd1-8944-9cb3b1e3fef2";
        //const string certificationPath = "C:\\Certs\\KBSchedule.lbriresorts.eu.pfx";
        //const string certificationPassword = "lbri_123!2";
        string mailFrom = Utils.getMailFrom(); 

        public  void sendMail(string subject, string body, string recipient)
        {
            X509Certificate2 certificate = new X509Certificate2(Utils.getCertificationPath(), Utils.getCertificationPassword());
            ClientCertificateCredential credential = new ClientCertificateCredential(
    tenantId: Utils.getTenantID(),
    clientId:  Utils.getClientID(),
    clientCertificate: certificate
);
            ConfidentialClientApplication app = (
            ConfidentialClientApplication)ConfidentialClientApplicationBuilder.Create(Utils.getClientID())
                .WithCertificate(certificate)
                .WithTenantId(Utils.getTenantID())
                .Build();
            var result = app.AcquireTokenForClient(scopes).ExecuteAsync().GetAwaiter().GetResult();
            string accessToken = GetAccessToken(app);
            body += Environment.NewLine + "<br/><hr/><br/><br/>" + Environment.NewLine + Environment.NewLine;
            sendGraphMail(subject, body, recipient, accessToken).GetAwaiter().GetResult();
            Console.WriteLine("Mail to recipient:" + recipient + " sent" ); 
            
        }

        private  string GetAccessToken(ConfidentialClientApplication app)
        {
            return GetAccessTokenAsync(app).GetAwaiter().GetResult();
        }

        private  async Task<string> GetAccessTokenAsync(ConfidentialClientApplication app)
        {
            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            return result.AccessToken;
        }


        private  async Task sendGraphMail(string mailSubject, string mailBody, string recipients, string accessToken)
        {
            var recipientList = recipients
            .Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(r => new
            {
                emailAddress = new
                {
                    address = r.Trim()
                }
            })
            .ToArray();

            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);

                var mail = new
                {
                    message = new
                    {
                        subject = mailSubject,
                        body = new
                        {
                            contentType = "HTML",
                            content = mailBody, 
                        },
                        toRecipients = recipientList
                    },
                    saveToSentItems = "true"
                };
                try
                {
                    var json = JsonConvert.SerializeObject(mail);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await httpClient.PostAsync(
                        "https://graph.microsoft.com/v1.0/users/" + mailFrom + "/sendMail",
                        content);
                    response.EnsureSuccessStatusCode();
                    Console.WriteLine("Mail sent to " + recipients);
                }
                catch (Exception ex) {
                    Console.WriteLine("Error (mail to:)" +recipients + " ==> " + ex.ToString()); 
                }
            }
        }

       
    }
}
