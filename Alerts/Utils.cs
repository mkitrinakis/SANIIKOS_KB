using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


namespace Alerts
{
    internal static class Utils
    {
        public static string baseUrl = "https://lbriresorts.sharepoint.com";
        public static string contentUploadSiteUrl = "https://lbriresorts.sharepoint.com/sites/KB-ContentUpload/";
        public static string[] getSiteUrls() { return ConfigurationManager.AppSettings["SiteCollections"].Split(';'); }
        public static string getAdminEmail() { return ConfigurationManager.AppSettings["AdminEmail"];}
        public static string getAdminEmailSubject() { return ConfigurationManager.AppSettings["AdminEmailSubject"]; }
        public static string getCertificationPassword() { return ConfigurationManager.AppSettings["CertificationPassword"]; }
        public static string getCertificationPath() { return ConfigurationManager.AppSettings["CertificationPath"]; }
        public static string getTenantID() { return ConfigurationManager.AppSettings["TenantID"]; }
        public static string getClientID() { return ConfigurationManager.AppSettings["ClientID"]; }
        public static int getThreshholdDays()
        {
            try
            {
                return Convert.ToInt32(ConfigurationManager.AppSettings["ThreshHoldDays"]);
            }
            catch { return 30; }
        }

        public static bool  getProductionMode() {
            string pMode = ConfigurationManager.AppSettings["ProductionMode"]; 
            if (pMode != null && (pMode.Equals("Yes", StringComparison.InvariantCultureIgnoreCase) || pMode.Equals("True", StringComparison.InvariantCultureIgnoreCase))) { return true; }
            return false;  }
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
                            "LibraryAdmins",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        publishersGroup = group;
                        break;
                    }
                }
            }

            if (publishersGroup == null)
            {
                Console.WriteLine("No LibraryAdmins group found.");
                return "";
            }

            UserCollection users = publishersGroup.Users;
            ctx.Load(users);
            ctx.ExecuteQuery();
            Console.WriteLine($"Group: {publishersGroup.Title}");
            return String.Join(";", users.Select(x => (x.Email ?? "")).ToList<string>());
        }
    }
}

