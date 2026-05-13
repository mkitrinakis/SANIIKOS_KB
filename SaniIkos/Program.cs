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
using Microsoft.IdentityModel.Tokens;

namespace SaniIkos
{
    class DocumentInfo
    {
        public string RootSiteTitle { get; set; }
        public string LibraryUrl { get; set; }
        public string LibraryTitle { get; set; }
        public string FileName { get; set; }
        public string FileUrl { get; set; }
        //public string DocumentName { get; set; }
        public string Title { get; set; }
        public int ID { get; set; }
        public DateTime? ValidUntil { get; set; }
        public string OwnerEmail { get; set; }
        public string OwnerName { get; set; }
        public string DispLibrary() { return "<a href=\"" + LibraryUrl + "\">" + LibraryTitle + "</a>"; }
        public string DispValidUntil() { return ValidUntil == null ? "N/A" : ((DateTime)ValidUntil).ToString("dd/MM/yyyy"); }

        public string DispTitle()
        {
            string itemUrl = LibraryUrl + "/Forms/DispForm.aspx?ID=" + ID.ToString();
            return "<a href=\"" + itemUrl + "\">" + Title + "</a>";
        }
        public string DispFileName() { return "<a href=\"" + FileUrl + "\">" + FileName + "</a>"; }
    }

    internal class Program
    {
        const string baseUrl = "https://lbriresorts.sharepoint.com";
        const string contentUploadSiteUrl = "https://lbriresorts.sharepoint.com/sites/KB-ContentUpload/";
        const string expirationDateFieldName = "Expiration_x0020_Date12";
        const string clientID = "XXXXXXXXXX";
        const string tenantID = "XXXXXXXXXXXXX";
        const string certificationPath = "XXXXXXXXXXX";
        const string certificationPassword = "XXXXXXXXXX";


        static void Main(string[] args)
        {
            
            string[] siteUrls = ConfigurationManager.AppSettings["SiteCollections"].Split(';');
            //var username = ConfigurationManager.AppSettings["Username"];
            //var password = ConfigurationManager.AppSettings["Password"];
            string adminEmail = ConfigurationManager.AppSettings["AdminEmail"];
            List<DocumentInfo> allResults = new List<DocumentInfo>();

            //AuthenticationManager authManager = AuthenticationManager.CreateWithCertificate(ClientID,  "C:\\Users\\mkitrinakis\\Desktop\\TestCert1.pfx", "abc123!2", tenantID);
            AuthenticationManager authManager = AuthenticationManager.CreateWithCertificate(clientID, certificationPath, certificationPassword, tenantID);
            {
                foreach (string siteUrl in siteUrls)
                {
                    ProcessSite(authManager, siteUrl, allResults);
                }

            }

            // Group by Owner
            var grouped = allResults
                .Where(x => x.OwnerEmail != null)
                .GroupBy(x => x.OwnerEmail);

            foreach (var group in grouped)
            {
                SendEmail(authManager, group.Key, group.ToList());
            }

            // Admin summary
         //   SendSummaryEmail(adminEmail, allResults);


            Console.WriteLine("Completed.");
        }


        static void ProcessSite(AuthenticationManager authManager, string siteUrl, List<DocumentInfo> allResults)
        {
            //string siteUrl = "https://intrrusttest.sharepoint.com/sites/MarkosCom1";
            using (var ctx = authManager.GetContext(siteUrl))
            {
                ctx.Load(ctx.Web, w => w.Title, w => w.Webs, w => w.Lists);
                ctx.ExecuteQuery();
                ProcessWeb(ctx, ctx.Web, ctx.Web.Title, allResults);
            }
        }

        static void ProcessWeb(ClientContext ctx, Web web, string rootTitle, List<DocumentInfo> results)
        {
            ctx.Load(web, w => w.Title, w => w.Webs, w => w.Lists.Include(l=>l.Title,  l=>l.RootFolder.ServerRelativeUrl));
            ctx.ExecuteQuery();
            foreach (var list in web.Lists)
            {
                if (list.Title.Contains("Private") || list.Title.Contains("Public"))
                {
                    ProcessLibrary(ctx, list, rootTitle, results);
                }
            }
            foreach (var subWeb in web.Webs)
            {
                ProcessWeb(ctx, subWeb, rootTitle, results);
            }
        }

        static void ProcessLibrary(ClientContext ctx, Microsoft.SharePoint.Client.List list, string rootTitle, List<DocumentInfo> results)
        {
            DateTime threshold = DateTime.Now.AddMonths(1);
            string viewXml = "<Lt><FieldRef Name='" + expirationDateFieldName + "'/><Value Type='DateTime'>" + threshold.ToString("s") + "Z</Value></Lt>";
            viewXml = "<Where>" + viewXml + "</Where>";
            viewXml = "<Query>" + viewXml + "</Query>";
            viewXml = "<View>" + viewXml + "</View>";
            CamlQuery query = new CamlQuery { ViewXml = viewXml };
            var items = list.GetItems(query);
            ctx.Load(items, i => i.Include(
                item => item["Title"],
                item => item["FileLeafRef"],
                item => item["FileRef"],
                item => item.Id,
                item => item[expirationDateFieldName],
                item => item["Owner"]
            ));
            ctx.ExecuteQuery();
            foreach (var item in items)
            {
                var owners = item["Owner"] as FieldUserValue[];
                var owner = owners != null && owners.Length > 0 ? owners[0] : null;
                results.Add(new DocumentInfo
                {
                    Title = item["Title"]?.ToString(),
                    FileName = item["FileLeafRef"]?.ToString(),
                    FileUrl = baseUrl + item["FileRef"]?.ToString(),
                    ID = item.Id,
                    LibraryTitle = list.Title,
                    LibraryUrl = baseUrl + list.RootFolder.ServerRelativeUrl,
                    RootSiteTitle = rootTitle,
                    OwnerEmail = owner?.Email,
                    OwnerName = owner?.LookupValue,
                    ValidUntil = (DateTime?)item["Expiration_x0020_Date12"]
                });
            }
        }

        static void SaveEmailToSharePoint(AuthenticationManager authManager, string to, string subject, string body)
        {
            using (var ctx = authManager.GetContext(contentUploadSiteUrl))
            {
                {
                    var list = ctx.Web.Lists.GetByTitle("MailHistory");

                    var itemCreate = new ListItemCreationInformation();
                    var item = list.AddItem(itemCreate);

                    item["Title"] = subject;
                    item["Body"] = body;
                    item["Recipient"] = to;

                    item.Update();
                    ctx.ExecuteQuery();
                    item.BreakRoleInheritance(copyRoleAssignments: false, clearSubscopes: true);
                    // Role = Read
                    var readRole = ctx.Web.RoleDefinitions.GetByType(Microsoft.SharePoint.Client.RoleType.Reader);
                    var readBindings = new RoleDefinitionBindingCollection(ctx) { readRole };
                    // Grant Read to SharePoint group "Auditors"
                    var auditors = ctx.Web.SiteGroups.GetByName("Auditors");
                    item.RoleAssignments.Add(auditors, readBindings);
                    // Grant Read to the recipient user (from email)
                    var recipientUser = ctx.Web.EnsureUser($"i:0#.f|membership|{to}");
                    item.RoleAssignments.Add(recipientUser, readBindings);
                    item.Update();
                    ctx.ExecuteQuery();
                }
            }
        }

        //static void SendEmail(string to, string body)
        //{
        //    var ctx = GetContext("https://tenant.sharepoint.com");

        //    var emailProps = new EmailProperties
        //    {
        //        To = new List<string> { to },
        //        Subject = "Documents Expiring Soon",
        //        Body = body
        //    };

        //    Utility.SendEmail(ctx, emailProps);
        //    ctx.ExecuteQuery();
        //}

        static void SendEmail(AuthenticationManager authManager, string to, List<DocumentInfo> docs)
        {
            using (var ctx = authManager.GetContext(contentUploadSiteUrl))
            {
                StringBuilder body = new StringBuilder();
                string subject = "Documents nearing expiration owned by " + to;
                body.Append("<h3>" + subject + "</h3>");
                body.AppendLine("<table>");
                body.AppendLine("<tr>");
                body.AppendLine("<th>Site</th><th>Library</th><th>ExpirationDate</th><th>Document Metadata</th><th>Document Body</th>");
                body.AppendLine("</tr>");
                foreach (var doc in docs)
                {
                    body.AppendLine("<tr>");
                    body.Append("<td>" + doc.RootSiteTitle + "</td>");
                    body.Append("<td>" + doc.DispLibrary() + "</td>");
                    body.Append("<td>" + doc.DispValidUntil() + "</td>");
                    body.Append("<td>" + doc.DispTitle() + "</td>");
                    body.Append("<td>" + doc.DispFileName() + "</td>");
                    body.AppendLine("</tr>");
                }
                body.AppendLine("</table>");

                var emailProps = new EmailProperties
                {
                    To = new List<string> { to },
                    Subject = "Documents Expiring Soon",
                    Body = body.ToString()
                };
                Console.WriteLine(" to log email:  " + body);
                SaveEmailToSharePoint(authManager, to, subject, body.ToString());
                Microsoft.SharePoint.Client.Utilities.Utility.SendEmail(ctx, emailProps);
                Console.WriteLine(" to send mail:  " + body);
                ctx.ExecuteQuery();
            }
        }

        
    }
}

    
