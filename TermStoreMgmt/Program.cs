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
using AuthenticationManager = PnP.Framework.AuthenticationManager;

using System.Security.Cryptography;


using static System.Net.WebRequestMethods;
using FieldUserValue = Microsoft.SharePoint.Client.FieldUserValue;

using ListItem = Microsoft.SharePoint.Client.ListItem;
using System.Collections;


namespace TermStoreMgmt
{
    internal class Program
    {
        static void Main(string[] args)
        {
            AuthenticationManager authManager = AuthenticationManager.CreateWithCertificate(Utils.getClientID(), Utils.getCertificationPath(), Utils.getCertificationPassword(), Utils.getTenantID());
            using (var ctx = authManager.GetContext(Utils.contentUploadSiteUrl))
            {
                ctx.Load(ctx.Web, w => w.Title, w => w.Webs, w => w.Lists);
                ctx.ExecuteQuery();
                SyncTermStore syncTermStore = new SyncTermStore();
                syncTermStore.run(ctx);
                //ImportTermStore importTermStore = new ImportTermStore();
                //importTermStore.run(ctx); 
            }

        }


    }
}
