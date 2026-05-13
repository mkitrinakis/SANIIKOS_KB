using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Runtime.Remoting.Contexts;

namespace TermStoreMgmt
{
    internal class ImportTermStore
    {
        public void run(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();

            TermGroup group = termStore.Groups.GetByName(Utils.termGroupName);
            TermSet termSet = group.TermSets.GetByName(Utils.termSetName);

            ctx.Load(termSet);
            ctx.Load(termSet.Terms);
            ctx.ExecuteQuery();

            List targetList = ctx.Web.Lists.GetByTitle(Utils.termListTitle);

            Console.WriteLine($"Reading TermSet: {Utils.termSetName}");
            Console.WriteLine();

            foreach (Term term in termSet.Terms)
            {
                ProcessTerm(ctx, targetList, term, "");
            }

            Console.WriteLine("Completed.");
        }



        private  void ProcessTerm(
            ClientContext context,
            List targetList,
            Term term,
            string parentPath)
        {
            context.Load(term);
            context.Load(term.Terms);
            context.ExecuteQuery();

            string currentPath = string.IsNullOrEmpty(parentPath)
                ? term.Name
                : $"{parentPath} - {term.Name}";

            Console.WriteLine($"Creating item: {currentPath}");

            // Create SharePoint List Item
            ListItemCreationInformation itemCreateInfo =
                new ListItemCreationInformation();

            ListItem item = targetList.AddItem(itemCreateInfo);

            // Title column
            item["Title"] = currentPath;

            // Custom column named 'key'
            item["TermID"] = term.Id.ToString();

            item.Update();

            context.ExecuteQuery();

            // Recursive processing for child terms
            foreach (Term childTerm in term.Terms)
            {
                ProcessTerm(context, targetList, childTerm, currentPath);
            }
        }
    }
}
