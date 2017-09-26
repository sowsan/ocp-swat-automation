using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace OCP.Swat.Automation
{
    class Program
    {
        private static ClientContext context;

        static void Main(string[] args)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your user name");
            Console.ForegroundColor = defaultForeground;
            string userName = Console.ReadLine();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your password.");
            Console.ForegroundColor = defaultForeground;
            SecureString password = GetPasswordFromConsoleInput();

            using (context = new ClientContext("<siteurl>"))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                context.Load(context.Web, w => w.Title);
                context.ExecuteQuery();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is: " + context.Web.Title);
                Console.ForegroundColor = defaultForeground;

                // Assume the web has a list named "Announcements". 
                List swarmList = context.Web.Lists.GetByTitle("Swarm Form");               


                // This creates a CamlQuery that has a RowLimit of 100, and also specifies Scope="RecursiveAll" 
                // so that it grabs all list items, regardless of the folder they are in. 
                CamlQuery query = CamlQuery.CreateAllItemsQuery(500);

                /* query.ViewXml = "<View><ViewFields><FieldRef Name='ID'/>" +
                    "<FieldRef Name='Title'/><FieldRef Name='Request Status'/>" +
                    "</ViewFields><RowLimit>5</RowLimit></View>";*/


                ListItemCollection swarmRequests = swarmList.GetItems(query);

                context.Load(
                swarmRequests,
                items => items.Take(500).Include(
                item => item["Title"],
                item => item["ID"],
                item => item["Partner"],
                item => item["Request_x0020_Status"],
                item => item["Swarm_x0020_Supporting_x0020_Arc"],
                item => item["Swat_x0020_Records_x0020_Generat"],
                item => item["Assigned_x0020_Swarm_x0020_Archi"]));

                // Retrieve all items in the ListItemCollection from List.GetItems(Query). 

                context.ExecuteQuery();

                Console.WriteLine("Total Items " + swarmRequests.Count.ToString());

                int count = 0;

                foreach (ListItem request in swarmRequests)
                {

                    if (request["Swat_x0020_Records_x0020_Generat"] != null && request["Request_x0020_Status"].ToString() == "Assigned")
                    {
                        if (request["Swat_x0020_Records_x0020_Generat"].ToString() == "No")
                        {
                            //Console.WriteLine(request["Swat_x0020_Records_x0020_Generat"].ToString() + " ## " + request["Request_x0020_Status"].ToString());
                           // request["Swat_x0020_Records_x0020_Generat"] = "No";
                           // request.Update();
                            count++;
                            CreateSwatRecords(request);
                        }
                        // Console.WriteLine(request["Swat_x0020_Records_x0020_Generat"].ToString() + " ## " + request["Request_x0020_Status"].ToString());
                        
                    }
                }



                context.ExecuteQuery();
                Console.WriteLine("Total swarm requests assigned with no swat records : " + count.ToString());
                Console.ReadLine();
            }

        }
        

        private static void CreateSwatRecords(ListItem request){


            List activityList = context.Web.Lists.GetByTitle("Activity");

            CamlQuery activityQuery = CamlQuery.CreateAllItemsQuery(50);

            ListItemCollection activityItems = activityList.GetItems(activityQuery);

            context.Load(
            activityItems,
            items => items.Take(50).Include(
            item => item["Title"],
            item => item["ID"],
            item => item["Auto_x002d_Generate"],
            item => item["Display_x0020_Value"],
            item => item["ID"]));

            // Retrieve all items in the ListItemCollection from List.GetItems(Query). 

            context.ExecuteQuery();

            Console.WriteLine("Total Activities " + activityItems.Count.ToString());

            int count = 0;                   
           
            foreach (var activity in activityItems)
            {
                if(Convert.ToBoolean(activity["Auto_x002d_Generate"]))
                {
                    List swatList = context.Web.Lists.GetByTitle("Swat Form");
                    ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                    ListItem swatActivity = swatList.AddItem(itemInfo);
                    swatActivity["Swarm_x0020_Request"] = request["ID"];
                    swatActivity["Title"] = activity["Display_x0020_Value"];
                    swatActivity["Activity_x0020_Type"] = activity["ID"];
                    swatActivity["Lead_x0020_Architect"] = request["Assigned_x0020_Swarm_x0020_Archi"];
                    swatActivity["Supporting_x0020_Architects"] = request["Swarm_x0020_Supporting_x0020_Arc"];

                    swatActivity.Update();
                }
            }
            request["Swat_x0020_Records_x0020_Generat"] = "Yes";
            request.Update();

            context.ExecuteQuery();           

            Console.WriteLine("Swarm Activities careted for Swarm Request : "+ request["ID"].ToString());


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
