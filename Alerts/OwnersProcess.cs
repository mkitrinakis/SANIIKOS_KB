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
using ListItem = Microsoft.SharePoint.Client.ListItem;
using System.Collections;
using System.Diagnostics.Eventing.Reader;

namespace Alerts
{
    internal class OwnersProcess
    {
        public static string expirationDateFieldName = "Expiration_x0020_Date12";
        public static List<string> allErrors = new List<string>();

        public void run()
        {
            AuthenticationManager authManager = AuthenticationManager.CreateWithCertificate(Utils.getClientID(), Utils.getCertificationPath(), Utils.getCertificationPassword(), Utils.getTenantID());
            runForRole(authManager, false); // for Owners 
            runForRole(authManager, true); // escalate to LibraryAdmins 
            if (allErrors.Count > 0 && (!String.IsNullOrEmpty(allErrors[0])))
            {
                string body = String.Join("<br/>" + Environment.NewLine, allErrors);
                MailUtils mailUtils = new MailUtils();
                mailUtils.sendMail(Utils.getAdminEmailSubject(), body, Utils.getAdminEmail());
                SaveEmailToSharePoint(authManager, Utils.getAdminEmail(), Utils.getAdminEmailSubject(), body);
            }
        }

        public void runForRole(AuthenticationManager authManager,  bool escalate) {
            List<DocumentInfo> allResults = new List<DocumentInfo>();
            List<DocumentInfo> allResultsEscalate = new List<DocumentInfo>();
            try
            {

                //MailUtils mailUtils = new MailUtils(); 
                //mailUtils.sendMail("Test subject", "test body", "mkitrinakis@inttrust.gr"); 
                //AuthenticationManager authManager = AuthenticationManager.CreateWithCertificate(ClientID,  "C:\\Users\\mkitrinakis\\Desktop\\TestCert1.pfx", "abc123!2", tenantID);
                

                foreach (string siteUrl in Utils.getSiteUrls())
                {
                    ProcessSite(authManager, siteUrl, allResults, escalate);
                }
                
                if (escalate)
                {
                var    grouped = allResults
                .GroupBy(x => x.LibraryAdmins);
                    foreach (var group in grouped)
                    {
                        SendEmail(authManager, group.Key, group.ToList(), escalate);
                    }
                }
                else
                {
                    var grouped = allResults
                   .Where(x => x.Owner != null)
                   .GroupBy(x => x.Owner.Email);
                    foreach (var group in grouped)
                    {
                        SendEmail(authManager, group.Key, group.ToList(), escalate);
                    }
                }

                
            }
            catch (Exception ex)
            {
                string errMsg = "ERROR on Main ==> " + ex.Message;
                Console.WriteLine(errMsg);
                allErrors.Add(errMsg);
            }

          


            Console.WriteLine("Alerts completed on:" + System.DateTime.Now.ToString());
        }


         void ProcessSite(AuthenticationManager authManager, string siteUrl, List<DocumentInfo> allResults, bool escalate)
        {
            Console.WriteLine("Processing Site: " + siteUrl);
            try
            {
                using (var ctx = authManager.GetContext(siteUrl))
                {
                    ctx.Load(ctx.Web, w => w.Title, w => w.Webs, w => w.Lists);
                    ctx.ExecuteQuery();
                    ProcessWeb(ctx, ctx.Web, ctx.Web.Title, allResults, escalate);
                }
            }
            catch (Exception ex)
            {
                string errMsg = "ERROR on ProcessSite :" + siteUrl + " ==> " + ex.Message;
                Console.WriteLine(errMsg);
                allErrors.Add(errMsg);
            }
        }

         void ProcessWeb(ClientContext ctx, Web web, string rootTitle, List<DocumentInfo> results, bool escalate)
        {
            try
            {
                string libraryAdmins = null;
                ctx.Load(web, w => w.Title, w => w.Webs, w => w.Lists.Include(l => l.Title, l => l.RootFolder.ServerRelativeUrl));
                ctx.ExecuteQuery();
                foreach (var list in web.Lists)
                {
                    if (list.Title.Contains("Private") || list.Title.Contains("Public"))
                    {
                        ProcessLibrary(ctx, web, list, rootTitle, results, escalate, ref libraryAdmins);
                    }
                }
                foreach (var subWeb in web.Webs)
                {
                    ProcessWeb(ctx, subWeb, rootTitle, results, escalate);
                }
            }
            catch (Exception ex)
            {
                string errMsg = "ERROR on ProcessWeb :" + (web?.Title ?? "N/A") + " ==> " + ex.Message;
                Console.WriteLine(errMsg);
                allErrors.Add(errMsg);
            }
        }

         void ProcessLibrary(ClientContext ctx, Web web, Microsoft.SharePoint.Client.List list, string rootTitle, List<DocumentInfo> results, bool escalate, ref string libraryAdmins)
        {
            try
            {
                DateTime threshold = escalate ? DateTime.Now.AddDays(Utils.getThreshholdDaysEscalate()) :  DateTime.Now.AddDays(Utils.getThreshholdDays());
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
                    if (libraryAdmins == null) { libraryAdmins = Utils.getLibraryAdmins(ctx, web); } // only first time get library admins 
                    var owners = item["Owner"] as FieldUserValue[];
                    var owner = owners != null && owners.Length > 0 ? owners[0] : null;
                    results.Add(new DocumentInfo
                    {
                        Title = item["Title"]?.ToString(),
                        FileName = item["FileLeafRef"]?.ToString(),
                        FileUrl = Utils.baseUrl + item["FileRef"]?.ToString(),
                        ID = item.Id,
                        LibraryTitle = list.Title,
                        LibraryUrl = Utils.baseUrl + list.RootFolder.ServerRelativeUrl,
                        LibraryAdmins = list.Title.Contains("Public") ? libraryAdmins : "",
                        RootSiteTitle = rootTitle,
                        Owner = owner,
                        OwnerEmail = owner?.Email,
                        OwnerName = owner?.LookupValue,
                        ValidUntil = (DateTime?)item["Expiration_x0020_Date12"]
                    });
                }
            }
            catch (Exception ex)
            {
                string errMsg = "ERROR on ProcessLibrary :" + (list?.Title ?? "N/A") + " ==> " + ex.Message;
                Console.WriteLine(errMsg);
                allErrors.Add(errMsg);
            }
        }

         void SendEmail(AuthenticationManager authManager, string receivers, List<DocumentInfo> docs, bool escalate)
        {
            try
            {
                if (docs != null && docs.Count > 0)
                {
                    FieldUserValue owner = (FieldUserValue)docs[0].Owner;
                    using (var ctx = authManager.GetContext(Utils.contentUploadSiteUrl))
                    {
                        StringBuilder body = new StringBuilder();
                        string subject =escalate ? "KB Documents nearing expiration in libraries of your administration" :  ("KB Documents nearing expiration owned by " + owner.LookupValue);
                        body.Append("<h3>" + subject + "</h3>");
                        body.AppendLine("<table>");
                        body.AppendLine("<tr>");
                        body.AppendLine("<th>Site</th><th>Library</th><th>ExpirationDate</th><th>Document Metadata</th><th>Document Body</th><th>Owner</th><th>Library Admins</th>");
                        body.AppendLine("</tr>");
                        foreach (var doc in docs)
                        {
                            body.AppendLine("<tr>");
                            body.Append("<td>" + doc.RootSiteTitle + "</td>");
                            body.Append("<td>" + doc.DispLibrary() + "</td>");
                            body.Append("<td>" + doc.DispValidUntil() + "</td>");
                            body.Append("<td>" + doc.DispTitle() + "</td>");
                            body.Append("<td>" + doc.DispFileName() + "</td>");
                            body.Append("<td>" + doc.Owner?.Email + "</td>");
                            body.Append("<td>" + doc.LibraryAdmins + "</td>");
                            body.AppendLine("</tr>");
                        }
                        body.AppendLine("</table>");
                        body.AppendLine("<br/>");
                        body.AppendLine("<br/>");
                        if (escalate)
                        {
                            body.AppendLine("This is an automated notification regarding all documents within the KB-libraries you are administrator,and that are assigned to owners with an expiration date falling within " + Utils.getThreshholdDaysEscalate() + " days from today. In addition, document owners receive notifications for all documents with an expiration date falling within " + Utils.getThreshholdDays() +  " days from today. <br/> Please note that this is an automated message.Kindly <b>do not reply</b> this message.");
                        }
                        else
                        {
                            body.AppendLine("This is an automated notification regarding all KB-libraries' documents you are assigned as 'Owner' with an expiration date falling within " + Utils.getThreshholdDays() + " days  from now.  <br/> Please note that this is an automated message.Kindly <b>do not reply</b> this message.");
                        }
                        //var emailProps = new EmailProperties
                        //{
                        //    To = new List<string> { to },
                        //    Subject = "Documents Expiring Soon",
                        //    Body = body.ToString()
                        //};
                        Console.WriteLine(" to log email:  " + body);
                        SaveEmailToSharePoint(authManager, receivers, subject, body.ToString());
                        if (Utils.getProductionMode() && !String.IsNullOrEmpty(receivers))
                        {
                            MailUtils mailUtils = new MailUtils();
                            mailUtils.sendMail(subject, body.ToString(), receivers); 
                            //mailUtils.sendMail(subject, body.ToString(), "mkitrinakis@inttrust.gr;markos.kitrinakis@gmail.com");
                        }
                        if (String.IsNullOrEmpty(receivers))
                        {
                            allErrors.Add("No recipients to send Alert mail (with subject:" + subject + " )"); 
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errMsg = "ERROR on SendEmail to " + receivers + " ==> " + ex.Message;
                Console.WriteLine(errMsg);
                allErrors.Add(errMsg);
            }
        }

         void SaveEmailToSharePoint(AuthenticationManager authManager, string recipients, string subject, string body)
        {
            try
            {
                using (var ctx = authManager.GetContext(Utils.contentUploadSiteUrl))
                {
                    {
                        var list = ctx.Web.Lists.GetByTitle("MailHistory");

                        ListItemCreationInformation itemCreate = new ListItemCreationInformation();
                        ListItem item = list.AddItem(itemCreate);
                        ctx.ExecuteQuery();
                        item["Title"] = subject;
                        item["Body"] = body;
                        item["MailTo"] = recipients; 
                        item.Update();
                        ctx.ExecuteQuery();
                        item.Update();
                        ctx.ExecuteQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                string errMsg = "ERROR on SaveEmailToSharePoint (" + subject + ") ==> " + ex.Message;
                Console.WriteLine(errMsg);
                allErrors.Add(errMsg);
            }
        }

        //static void SaveEmailToSharePoint_OldNotUsed(AuthenticationManager authManager, FieldUserValue owner, string subject, string body)
        //{
        //    try
        //    {
        //        using (var ctx = authManager.GetContext(Utils.contentUploadSiteUrl))
        //        {
        //            {
        //                var list = ctx.Web.Lists.GetByTitle("MailHistory");

        //                ListItemCreationInformation itemCreate = new ListItemCreationInformation();
        //                ListItem item = list.AddItem(itemCreate);
        //                Microsoft.SharePoint.Client.User recipientUser = ctx.Web.EnsureUser($"i:0#.f|membership|{owner.LookupValue}");
        //                ctx.Load(recipientUser);
        //                ctx.Load(recipientUser, x => x.Email);
        //                ctx.ExecuteQuery();
        //                item["Title"] = subject;
        //                item["Body"] = body;
        //                //Microsoft.SharePoint.Client.User recipient = ctx.Web.EnsureUser(owner.Email); 
        //                //ctx.Load(recipient);
        //                //ctx.ExecuteQuery(); 
        //                // item["Recipient"] = recipient;
        //                // item["Recipient"] = new FieldUserValue {  LookupId = owner.LookupId };
        //                item["MailTo"] = recipientUser?.Email;

        //                // item["Recipient"] = ctx.Web.EnsureUser($"i:0#.f|membership|{owner.LookupValue}");
        //                item.Update();
        //                ctx.ExecuteQuery();
        //                item.BreakRoleInheritance(copyRoleAssignments: false, clearSubscopes: true);
        //                // Role = Read
        //                var readRole = ctx.Web.RoleDefinitions.GetByType(Microsoft.SharePoint.Client.RoleType.Reader);
        //                var readBindings = new RoleDefinitionBindingCollection(ctx) { readRole };
        //                // Grant Read to SharePoint group "Auditors"
        //                var auditors = ctx.Web.SiteGroups.GetByName("Auditors");
        //                item.RoleAssignments.Add(auditors, readBindings);
        //                // Grant Read to the recipient user (from email)

        //                //var recipientUser = ctx.Web.EnsureUser(owner.);
        //                if (recipientUser != null) { item.RoleAssignments.Add(recipientUser, readBindings); }
        //                item.Update();
        //                ctx.ExecuteQuery();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        string errMsg = "ERROR on SaveEmailToSharePoint (" + subject + ") ==> " + ex.Message;
        //        Console.WriteLine(errMsg);
        //        allErrors.Add(errMsg);
        //    }
        //}
    }
}
