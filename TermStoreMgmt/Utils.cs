using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Marketplace.CorporateCuratedGallery;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace TermStoreMgmt
{
    internal static class Utils
    {
        public static string baseUrl = "https://lbriresorts.sharepoint.com";
        public static string contentUploadSiteUrl = "https://lbriresorts.sharepoint.com/sites/KB-ContentUpload/";
        public static string termGroupName = "KnB";
        public static string termSetName = "Hierarchy";
        public static string termListTitle = "TermStore-Hierarchy";
        
        public static string getCertificationPassword() { return ConfigurationManager.AppSettings["CertificationPassword"]; }
        public static string getCertificationPath() { return ConfigurationManager.AppSettings["CertificationPath"]; }
        public static string getTenantID() { return ConfigurationManager.AppSettings["TenantID"]; }
        public static string getClientID() { return ConfigurationManager.AppSettings["ClientID"]; }
        


        public static string getLibraryAdmins(ClientContext ctx, Web web)
        {

            // Load role assignments
            RoleAssignmentCollection roleAssignments = web.RoleAssignments;

            ctx.Load(
                roleAssignments,
                ras => ras.Include(
                    ra => ra.Member,
                    ra => ra.RoleDefinitionBindings));

            ctx.ExecuteQuery();

            Group publishersGroup = null;

            // Find first SP group containing "Publishers"
            foreach (RoleAssignment ra in roleAssignments)
            {
                if (ra.Member.PrincipalType == PrincipalType.SharePointGroup)
                {
                    Group group = ctx.Web.SiteGroups.GetById(ra.Member.Id);

                    ctx.Load(group);
                    ctx.ExecuteQuery();

                    if (group.Title.IndexOf(
                            "Publishers",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        publishersGroup = group;
                        break;
                    }
                }
            }

            if (publishersGroup == null)
            {
                Console.WriteLine("No Publishers group found.");
                return "";
            }

            // Load users of the group
            UserCollection users = publishersGroup.Users;

            ctx.Load(users);
            ctx.ExecuteQuery();

            Console.WriteLine($"Group: {publishersGroup.Title}");
            return String.Join(";", users.Select(x => (x.Email ?? "")).ToList<string>());
            //foreach (User user in users)
            //{
            //    Console.WriteLine(
            //        $"{user.Title} | {user.Email} | {user.LoginName}");
            //}
        }
    }
}

