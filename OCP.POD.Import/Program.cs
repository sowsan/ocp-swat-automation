using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using System.Security;
using System.IO;
using SP = Microsoft.SharePoint.Client;
using Microsoft.VisualBasic.FileIO;

namespace OCP.POD.Import
{
    class Program
    {
        private static ClientContext context;

        enum BooleanAliases
        {
            Yes = 1,            
            No = 0
        }

        static void Main(string[] args)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;
            List<SwarmRequest> swarmRequests = new List<SwarmRequest>();
            using (TextFieldParser parser = new TextFieldParser(@"test.csv"))
            {
                parser.Delimiters = new string[] { "," };
                while (true)
                {
                    string[] a = parser.ReadFields();
                    if (a == null)
                    {
                        break;
                    }

                    SwarmRequest swarmRequest = new SwarmRequest(){
                        Opportunity = a[0],
                        Partner = a[1],
                        PartnerType = a[2],
                        PDM = a[3],
                        PTS = a[4],
                        Manager = a[5],
                        Location = a[6],
                        IsVirtual = a[7],
                        PrimaryTechnology = a[8],
                        IsCompete = a[9],
                        InMarketDate = a[10],
                        Description = a[11]
                    };

                    swarmRequests.Add(swarmRequest);
                                   
                }
            }

            swarmRequests.RemoveAt(0);
    
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your user name");
            Console.ForegroundColor = defaultForeground;
            string userName = Console.ReadLine();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter your password.");
            Console.ForegroundColor = defaultForeground;
            SecureString password = GetPasswordFromConsoleInput();

            using (context = new ClientContext("https://url"))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                Web currentWeb = context.Web;
                context.Load(currentWeb);
                context.ExecuteQuery();

                SP.List oList = context.Web.Lists.GetByTitle("Swarm Form");

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                int updated = 0;

                foreach (SwarmRequest request in swarmRequests)
                {
                    User ptsUser = currentWeb.EnsureUser(request.PTS);
                    context.Load(ptsUser);

                    User pdmUser = currentWeb.EnsureUser(request.PDM);
                    context.Load(pdmUser);
                    
                    User managerUser = currentWeb.EnsureUser(request.Manager);
                    context.Load(managerUser);

                    context.ExecuteQuery();

                    FieldUserValue PTS = new FieldUserValue();
                    PTS.LookupId = ptsUser.Id;                   
               

                    FieldUserValue PDM = new FieldUserValue();
                    PDM.LookupId = pdmUser.Id;  

                    FieldUserValue Manager = new FieldUserValue();
                    Manager.LookupId = managerUser.Id;

                    ListItem oListItem = oList.AddItem(itemCreateInfo);
                    oListItem["Title"] = request.Opportunity + "-" + request.Partner;
                    oListItem["Partner"] = request.Partner;
                    oListItem["Partner_x0020_Type"] = request.PartnerType;
                    oListItem["PTS"] = PDM;
                    oListItem["PTS0"] = PTS;
                    oListItem["Assigned_x0020_Swarm_x0020_Archi"] = PTS;
                    oListItem["Request_x0020_Status"] = "Pod Assigned";
                    oListItem["Location"] = request.Location;
                    oListItem["Virtual_x0020_Meeting"] = request.IsVirtual;
                    oListItem["Workload"] = request.PrimaryTechnology;
                    oListItem["Compete"] = request.IsCompete;
                    oListItem["Scorecard_x0020_Start_x0020_Date"] = request.InMarketDate;
                    oListItem["Opportunity_x0020_Description"] = request.Description;
                    oListItem["Pod_x0020_Manager"] = Manager;
                    //oListItem["TXT_Lead_Architect"] = ptsUser.LoginName;
                    //oListItem["TXT_PTS_TE"] = ptsUser.n;
                    oListItem["TXT_Request_Status"] = "Pod Assigned";

                    oListItem.Update();

                    context.ExecuteQuery();

                    updated++;

                    Console.WriteLine("updated : " + updated.ToString());

                }

                Console.WriteLine("Total number of records imported : " + updated.ToString());
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
