using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Remoting.Contexts;

namespace TermStoreMgmt
{
    internal class SyncTermStore
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

            Console.WriteLine("Loading term store...");
            Dictionary<string, string> termDictionary = LoadAllTerms(ctx);
            if (termDictionary.Count < 10)
            {
                Console.WriteLine("Less than 10 terms found. Probably an error. Operation Aborted!"); 
                return; 
            }
            Console.WriteLine($"Loaded {termDictionary.Count} terms.");
            Console.WriteLine("Loading SharePoint list items...");
            List targetList = ctx.Web.Lists.GetByTitle(Utils.termListTitle);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = targetList.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQuery();

            Console.WriteLine($"Loaded {items.Count} list items.");

            // ============================================
            // BUILD LIST DICTIONARY
            // key = term GUID
            // value = ListItem
            // ============================================
            Dictionary<string, ListItem> listDictionary =
                new Dictionary<string, ListItem>();

            foreach (ListItem item in items)
            {
                if (item["TermID"] != null)
                {
                    string key = item["TermID"].ToString();

                    if (!listDictionary.ContainsKey(key))
                    {
                        listDictionary.Add(key, item);
                    }
                }
            }

            // ============================================
            // SYNCHRONIZE
            // ============================================

            int inserted = 0;
            int updated = 0;
            int deleted = 0;

            // --------------------------------------------
            // INSERT / UPDATE
            // --------------------------------------------
            foreach (var termEntry in termDictionary)
            {
                string termId = termEntry.Key;
                string desiredTitle = termEntry.Value;

                if (listDictionary.ContainsKey(termId))
                {
                    // Exists -> check update
                    ListItem existingItem = listDictionary[termId];

                    string currentTitle =
                        existingItem["Title"]?.ToString() ?? "";

                    if (currentTitle != desiredTitle)
                    {
                        Console.WriteLine(
                            $"UPDATE: {currentTitle} -> {desiredTitle}");

                        existingItem["Title"] = desiredTitle;
                        existingItem.Update();

                        updated++;
                    }
                }
                else
                {
                    // Insert
                    Console.WriteLine($"INSERT: {desiredTitle}");

                    ListItemCreationInformation itemCreateInfo =
                        new ListItemCreationInformation();

                    ListItem newItem =
                        targetList.AddItem(itemCreateInfo);

                    newItem["Title"] = desiredTitle;
                    newItem["TermID"] = termId;

                    newItem.Update();

                    inserted++;
                }
            }

            // --------------------------------------------
            // DELETE
            // --------------------------------------------
            HashSet<string> termIds =
                new HashSet<string>(termDictionary.Keys);

            foreach (var listEntry in listDictionary)
            {
                string itemKey = listEntry.Key;

                if (!termIds.Contains(itemKey))
                {
                    ListItem itemToDelete = listEntry.Value;

                    Console.WriteLine(
                        $"DELETE: {itemToDelete["Title"]}");

                    itemToDelete.DeleteObject();

                    deleted++;
                }
            }

            // Commit all changes in one batch
            ctx.ExecuteQuery();

            Console.WriteLine();
            Console.WriteLine("=================================");
            Console.WriteLine("Synchronization Completed");
            Console.WriteLine("=================================");
            Console.WriteLine($"Inserted : {inserted}");
            Console.WriteLine($"Updated  : {updated}");
            Console.WriteLine($"Deleted  : {deleted}");
        }


        private static Dictionary<string, string> LoadAllTerms(
           ClientContext context)
        {
            Dictionary<string, string> terms =
                new Dictionary<string, string>();

            TaxonomySession taxonomySession =
                TaxonomySession.GetTaxonomySession(context);

            TermStore termStore =taxonomySession.GetDefaultSiteCollectionTermStore();

            TermGroup group =termStore.Groups.GetByName(Utils.termGroupName);

            TermSet termSet =group.TermSets.GetByName(Utils.termSetName);

            context.Load(termSet);
            context.Load(termSet.Terms);
            context.ExecuteQuery();

            foreach (Term term in termSet.Terms)
            {
                ReadTermRecursive(
                    context,
                    term,
                    "",
                    terms);
            }

            return terms;
        }

        private static void ReadTermRecursive(
            ClientContext context,
            Term term,
            string parentPath,
            Dictionary<string, string> terms)
        {
            context.Load(term);
            context.Load(
       term,
       t => t.Name,
       t => t.Id,
       t => t.IsDeprecated,
       t => t.Terms);
            //context.Load(term.Terms);
            context.ExecuteQuery();
            if (term.IsDeprecated)
            {
                Console.WriteLine(
                    $"SKIP DEPRECATED: {term.Name}");

                return;
            }
            string currentPath =
                string.IsNullOrEmpty(parentPath)
                    ? term.Name
                    : $"{parentPath} - {term.Name}";

            string termId = term.Id.ToString();

            if (!terms.ContainsKey(termId))
            {
                terms.Add(termId, currentPath);
            }

            foreach (Term childTerm in term.Terms)
            {
                ReadTermRecursive(
                    context,
                    childTerm,
                    currentPath,
                    terms);
            }
        }
    }
}
