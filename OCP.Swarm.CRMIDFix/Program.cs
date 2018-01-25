using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace OCP.Swarm.CRMIDFix
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
                context.Load(context.Web, w => w.Title);
                context.ExecuteQuery();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is: " + context.Web.Title);
             
                // Assume the web has a list named "Announcements". 
                List swarmList = context.Web.Lists.GetByTitle("Swarm Form");               


                // This creates a CamlQuery that has a RowLimit of 100, and also specifies Scope="RecursiveAll" 
                // so that it grabs all list items, regardless of the folder they are in. 
               CamlQuery query = CamlQuery.CreateAllItemsQuery(2400);

                ListItemCollection swarmRequests = swarmList.GetItems(query);

                context.Load(
                swarmRequests,
                items => items.Take(2400).Include(
                item => item["Title"],
                item => item["ID"],
                item => item["CRM_x0020_Account_x0020_ID"]));

                // Retrieve all items in the ListItemCollection from List.GetItems(Query). 

                context.ExecuteQuery();

                Console.WriteLine("Total Items " + swarmRequests.Count.ToString());

                int count = 0;
                string crmId = string.Empty;

                foreach (ListItem request in swarmRequests)
                {
                    crmId = request["CRM_x0020_Account_x0020_ID"].ToString();
                    if(crmId.Contains(" "))
                    {
                        Console.WriteLine(request["ID"].ToString() + " "+  request["Title"].ToString());
                        crmId = crmId.Trim();
                        //request["CRM_x0020_Account_x0020_ID"] = crmId;
                        //request.Update();
                        //context.ExecuteQuery();
                        count++;
                    }
                }
                Console.WriteLine("Total " + count.ToString());
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
