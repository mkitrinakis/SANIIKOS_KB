using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace Alerts
{
    internal class DocumentInfo
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
        public FieldUserValue Owner { get; set; }
        public string DispLibrary() { return "<a href=\"" + LibraryUrl + "\">" + LibraryTitle + "</a>"; }
        public string DispValidUntil() { return ValidUntil == null ? "N/A" : ((DateTime)ValidUntil).ToString("dd/MM/yyyy"); }
        public string LibraryAdmins { get; set; }
        public string DispTitle()
        {
            string itemUrl = LibraryUrl + "/Forms/DispForm.aspx?ID=" + ID.ToString();
            return "<a href=\"" + itemUrl + "\">" + Title + "</a>";
        }
        public string DispFileName() { return "<a href=\"" + FileUrl + "\">" + FileName + "</a>"; }
    }
}
