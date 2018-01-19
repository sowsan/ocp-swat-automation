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

            using (context = new ClientContext("https://microsoft.sharepoint.com/teams/USDXISVCJ"))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                context.Load(context.Web, w => w.Title);
                context.ExecuteQuery();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is: " + context.Web.Title);
                Console.ForegroundColor = defaultForeground;
                UpdateFilterFields(context);
                /*
                // Assume the web has a list named "Announcements". 
                List swarmList = context.Web.Lists.GetByTitle("Swarm Form");               


                // This creates a CamlQuery that has a RowLimit of 100, and also specifies Scope="RecursiveAll" 
                // so that it grabs all list items, regardless of the folder they are in. 
                CamlQuery query = CamlQuery.CreateAllItemsQuery(500);

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
                */
                // context.ExecuteQuery();
                //  Console.WriteLine("Total swarm requests assigned with no swat records : " + count.ToString());
                Console.ReadLine();
            }

        }
        

        private static void UpdateSwarmFilterFields(ClientContext context)
        {
            try
            {
                List swarmList = context.Web.Lists.GetByTitle("Swarm Form");
              
                CamlQuery query = CamlQuery.CreateAllItemsQuery(2000);    

                ListItemCollection swarmItems = swarmList.GetItems(query);

                context.Load(
                swarmItems,
                items => items.Take(2000).Include(
                item => item["Request_x0020_Status"],
                item => item["Swat_x0020_Records_x0020_Generat"],
                item => item["Scorecard_x0020_Start_x0020_Date"]              
              ));

                context.ExecuteQuery();


                int updated = 0;

                int inMarketDateNull = 0;

                int inMarketDateNotNull = 0;


                foreach (ListItem swarmItem in swarmItems)
                {
                    if (swarmItem["Request_x0020_Status"] == "Assigned" && swarmItem["Swat_x0020_Records_x0020_Generat"] != "Yes"){
                        updated++;
                    }
                    //swarmItem["Scorecard_x0020_Start_x0020_Date"] = swarmItem["Scorecard_x0020_Start_x0020_Date"];
                    //swarmItem.Update();
                    //context.ExecuteQuery();
                    //updated++;
                    Console.WriteLine("updated " + updated.ToString());
                    //if (swarmItem["Scorecard_x0020_Start_x0020_Date"] == null)
                    //{
                    //    inMarketDateNull++;
                    //    //swarmItem["Scorecard_x0020_Start_x0020_Date"] = "In Market Date Unknown";
                    //    //swarmItem.Update();
                    //    //context.ExecuteQuery();
                    //    updated++;
                    //}
                    //else
                    //{
                    //    inMarketDateNotNull++;
                    //}

                   

                    // Console.WriteLine(swatItem["ID"].ToString() + " **hidden is**  " + ((FieldLookupValue)swatItem["Swarm_x0020_Request"]).LookupValue + " **hidden is** " + swatItem["TXT_Swarm_Request"].ToString());
                }

                Console.WriteLine("Swat acitivities not generated items " + updated.ToString());



                //   Console.WriteLine("inMarketDateNull : " + inMarketDateNull.ToString());
                //  Console.WriteLine("inMarketDateNotNull : "+ inMarketDateNotNull.ToString());

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private static void UpdateFilterFields(ClientContext context)
        {
            string TXT_Lead_Archiect = string.Empty;
            string TXT_PTS_TE = string.Empty;
            string TXT_Request_Status = string.Empty;
            string TXT_Supp_Arch1 = string.Empty;
            string TXT_Supp_Arch2 = string.Empty;
            string TXT_Supp_Arch3 = string.Empty;
            string TXT_Supp_Arch4 = string.Empty;

            try
            {
                List swarmList = context.Web.Lists.GetByTitle("Swarm Form");
                
                CamlQuery query = CamlQuery.CreateAllItemsQuery(3000);

                ListItemCollection swatItems = swarmList.GetItems(query);

                /*  context.Load(
                  swatItems,
                  items => items.Take(3000).Include(
                  item => item["Title"],
                  item => item["ID"],
                  item => item["TXT_Lead_Architect"],
                  item => item["Swarm_x0020_Request"],
                  item => item["TXT_Swarm_Request"],
                  item => item["Lead_x0020_Architect"],
                  item => item["Supporting_x0020_Architects"]
                ));*/


                context.Load(
                swatItems,
                items => items.Take(2000).Include(
                item => item["Title"],
                item => item["ID"],
                item => item["TXT_Lead_Architect"],
                item => item["TXT_PTS_TE"],
                item => item["PTS0"],
                item => item["Assigned_x0020_Swarm_x0020_Archi"],
                item => item["Swarm_x0020_Supporting_x0020_Arc"],
                item => item["Request_x0020_Status"]
              ));

                // Retrieve all items in the ListItemCollection from List.GetItems(Query). 

                context.ExecuteQuery();

                List<string> textFields = new List<string>();
                textFields.Add("TXT_Supp_Arch1");
                textFields.Add("TXT_Supp_Arch2");
                textFields.Add("TXT_Supp_Arch3");
                textFields.Add("TXT_Supp_Arch4");

                Console.WriteLine("Total Items " + swatItems.Count.ToString());

                int updated = 0;
                int morethan4Architects = 0;

                foreach (ListItem swatItem in swatItems)
                {
                    //if(swatItem["Request_x0020_Status"] == "Pod Assigned")
                    //{

                    //}
                    if(swatItem["Swarm_x0020_Supporting_x0020_Arc"] != null)
                    {
                        FieldUserValue[] architects = swatItem["Swarm_x0020_Supporting_x0020_Arc"] as FieldUserValue[];

                        int count = architects.ToArray().Count();

                        for (int i = 0; i < count; i++)
                        {
                            if (i <= 3)
                            {                             
                                swatItem[textFields[i]] = architects[i].LookupValue;
                            }
                               
                            else
                            {
                                morethan4Architects++;
                                Console.WriteLine("There are " + count.ToString() + "support architects for this item " + swatItem["ID"].ToString());
                            }
                                
                        }
                    }

                    if (swatItem["Assigned_x0020_Swarm_x0020_Archi"] != null)
                    {
                        swatItem["TXT_Lead_Architect"] = ((FieldUserValue)swatItem["Assigned_x0020_Swarm_x0020_Archi"]).LookupValue;       
                    }

                    if(swatItem["PTS0"] != null)
                    {
                        swatItem["TXT_PTS_TE"] = ((FieldUserValue)swatItem["PTS0"]).LookupValue;
                    }

                    /*
                     * swat activity
                    
                    if (swatItem["TXT_Lead_Architect"] == null && swatItem["TXT_Swarm_Request"] == null)
                      {

                        //swatItem["TXT_Lead_Architect"] = ((FieldLookupValue)swatItem["Lead_x0020_Architect"]).LookupValue;
                       // swatItem["TXT_Swarm_Request"] = ((FieldLookupValue)swatItem["Swarm_x0020_Request"]).LookupValue;
                       // Console.WriteLine(((FieldLookupValue)swatItem["Swarm_x0020_Request"]).LookupValue);
                      } */

                    swatItem["TXT_Request_Status"] = swatItem["Request_x0020_Status"];

                    swatItem.Update();
                    context.ExecuteQuery();
                    updated++;
                     Console.WriteLine("updated " + updated.ToString());

                   // Console.WriteLine(swatItem["ID"].ToString() + " **hidden is**  " + ((FieldLookupValue)swatItem["Swarm_x0020_Request"]).LookupValue + " **hidden is** " + swatItem["TXT_Swarm_Request"].ToString());
                }

                Console.WriteLine("There are total "+ morethan4Architects.ToString() + " has more than 4 supporting architcts");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);      
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
