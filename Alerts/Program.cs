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
  

    internal class Program
    {


        
        static void Main(string[] args)
        {
            Console.WriteLine ("Alerts starting on:" + System.DateTime.Now.ToString());
            if (Utils.getProductionMode())
            {
                Console.WriteLine("on PRODUCTION Mode (mails will be sent)");
            }
            else
            {
                Console.WriteLine("on TEST Mode (NO mails will be sent)");
            }
            OwnersProcess ownersProcess = new OwnersProcess();
            ownersProcess.run(); 
           
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

      


    }
}


