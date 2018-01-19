using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Security;
using System.IO;
using SP = Microsoft.SharePoint.Client;


namespace OCP.Workload.Export
{
    class Program
    {
        private static ClientContext context;

        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your user name");
           // Console.ForegroundColor = defaultForeground;
            string userName = Console.ReadLine();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your password.");
           // Console.ForegroundColor = defaultForeground;
            SecureString password = GetPasswordFromConsoleInput();

            using (context = new ClientContext("https://microsoft.sharepoint.com/teams/USDXISVCJ/"))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                Web currentWeb = context.Web;
                context.Load(currentWeb);
                context.ExecuteQuery();

                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(currentWeb.Context);

                TermStoreCollection termStores = taxonomySession.TermStores;

                currentWeb.Context.Load(termStores);

                currentWeb.Context.ExecuteQuery();

                TermStore termStore = termStores[0];

                currentWeb.Context.Load(termStore);

                currentWeb.Context.ExecuteQuery();

                TermGroupCollection termGroups = termStore.Groups;

                currentWeb.Context.Load(termGroups);

                currentWeb.Context.ExecuteQuery();

                TermGroup group = termGroups.GetByName("Site Collection - microsoft.sharepoint.com-teams-USDXISVCJ");

                TermSetCollection termSets = group.TermSets;

                currentWeb.Context.Load(termSets);

                currentWeb.Context.ExecuteQuery();

                TermCollection terms = termSets.GetByName("Workload").Terms;

                currentWeb.Context.Load(terms);

                currentWeb.Context.ExecuteQuery();

                Console.WriteLine(terms.Count);

                foreach (Term term in terms)
                {
                    Console.WriteLine(term.Name);
                }


            }
        }

        private static SecureString GetPasswordFromConsoleInput()
        {
            ConsoleKeyInfo info;

            //Get the user's password as a SecureString
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}
