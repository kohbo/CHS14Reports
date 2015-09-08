/*
 * Reporting application to parse PK notices
 * Written By: Juan Menendez
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Reflection;
using System.Globalization;
using System.IO;

namespace PKReporter
{
    public class Reporter
    {
        Outlook.Application OutlookApp;
        Outlook.Store PKStore;
        Outlook.Folder PKAlerts;
        ExcelCharter charter;

        private List<string[]> failures
        {
            set;
            get;
        }
        public List<string[]> Failures
        {
            get
            {
                return failures;
            }
        }


        public Reporter()
        {
            Console.ForegroundColor = ConsoleColor.Black;
            Console.BackgroundColor = ConsoleColor.White;
            Console.WriteLine("Starting PK Reporter\nWritten by Juan Menendez\n\n");
            Console.ResetColor();

            OutlookApp = GetApplicationObject();
            PKStore = OutlookApp.Session.GetStoreFromID(GetStoreID(OutlookApp.Session.Stores));
            PKAlerts = (Outlook.Folder)PKStore.GetRootFolder();
        }

        private void GenerateMailReport()
        {
            Console.WriteLine("Creating Report Email...");

            StringBuilder HTMLBody = new StringBuilder();
            HTMLBody.Append("<style>body{font-size:.8em;}table{border-collapse:collapse}td,th{text-align: center; padding: 2px 5px;}th{background: #ffa900; color:white;}.total{background:#3CC}#bysite>tr{background:black}</style>");
            HTMLBody.Append("<h2>PK Failure Notice Report</h2>\n");
            HTMLBody.AppendFormat("<p>{0} failure notices received since {1}\n", PKAlerts.Items.Count, PKAlerts.Items.GetFirst().ReceivedTime.ToString());
            HTMLBody.Append("<h3>Failures By Type</h3>\n\n");
            //HTMLBody.Append(GenerateBySiteCountReport());
            HTMLBody.Append(GenerateSiteByWeekReport());
            HTMLBody.Append(GenerateBySiteReport());

            Outlook.MailItem mailItem = OutlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "PK Failure Report as of " + System.DateTime.Now.ToString();
            mailItem.To = "CHS14_Service_Desk@chs14.net";
            mailItem.HTMLBody = HTMLBody.ToString();
            charter = new ExcelCharter(Failures);            
            try
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = System.ConsoleColor.Black;
                Console.WriteLine("Mail item created. Click send in Outlook to distribute report.");
                Console.ResetColor();
                mailItem.Display(true);
            }
            catch (COMException exc)
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.ForegroundColor = System.ConsoleColor.White;
                Console.WriteLine("Exception Occured. Ensure all dialog boxes are closed and try again.\n" + exc.Message);
                Console.ResetColor();
                Console.ReadKey();
            }
            finally
            {
                Marshal.ReleaseComObject(OutlookApp);
                OutlookApp = null;
            }
        }

        private void LoadEmailsToArray()
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            failures = new List<string[]>();
            Console.WriteLine("Collecting Data from Emails...");

            foreach (Outlook.MailItem item in PKAlerts.Items)
            {
                item.Categories = "Policed";
                string[] sourcebody = item.Body.Split('\n');
                string[] data = new string[9];
                /*
                 * 0: Date
                 * 1: Environment
                 * 2: Patient Location
                 * 3: Facility
                 * 4: Destination Group
                 * 5: Destination
                 * 6: Address
                 * 7: Details
                 * 8: Week
                 */
                data[0] = item.ReceivedTime.ToString();
                for (int index = 1; index < sourcebody.Length - 1; index++)
                {
                    data[index] = sourcebody[index - 1].Split(':')[1].Trim();
                }
                data[8] = Convert.ToString(cal.GetWeekOfYear(item.ReceivedTime, dfi.CalendarWeekRule, dfi.FirstDayOfWeek));
                failures.Add(data);
                
            }
        }

        Outlook.Application GetApplicationObject()
        {

            Outlook.Application application = null;

            // Check if there is an Outlook process running. 
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {

                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object. 
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            else
            {

                // If not, create a new instance of Outlook and log on to the default profile. 
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "", Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object. 
            return application;
        }

        private string GetStoreID(Outlook.Stores stores)
        {
            //Console.WriteLine("Enter the number corresponding with the data file holding PK alerts...");
            //int index = 0;
            //foreach(Outlook.Store store in stores)
            //{
            //    Console.WriteLine(index + ": " + store.DisplayName);
            //    index++;
            //}
            //Console.Write("Selection > ");
            //int sel = Convert.ToInt32(Console.Read()) - 47;
            //return stores[sel].StoreID;

            foreach (Outlook.Store store in stores)
            {
                if (store.DisplayName == "PK_Alerts")
                {
                    return store.StoreID;
                }
            }

            return "";
        }

        private void ReportToFile(StringBuilder HTMLString)
        {
            StreamWriter StreamOut = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "Report.html");
            StreamOut.Write(HTMLString.ToString());
            StreamOut.Close();
        }

        private void EnumerateStores(Outlook.Stores stores)
        {
            foreach (Outlook.Store store in stores)
            {
                if (store.IsDataFileStore == true)
                {
                    Debug.WriteLine(String.Format("Store: "
                    + store.StoreID
                    + "\n" + "File Path: "
                    + store.FilePath + "\n"));
                }
            }
        }

        private string GenerateSiteByWeekReport()
        {
            Console.WriteLine("Generating By Week Count Report...");
            StringBuilder HTMLDataString = new StringBuilder();
            StringBuilder HeaderString = new StringBuilder("<h3>Failures By Week</h3><table id='byweek'><tr><th></th>");
            StringBuilder SecHeaderString = new StringBuilder("<tr><th>All Total</th>");
            List<string> facilities = new List<string>();
            SortedDictionary<int, List<string>> WeeksWithDates = new SortedDictionary<int, List<string>>();
            foreach (string[] failure in failures)
            {
                if (!facilities.Contains(failure[3]))
                {
                    facilities.Add(failure[3]);
                }
                int weeknum = Convert.ToInt32(failure[8]);
                if (!WeeksWithDates.ContainsKey(weeknum))
                {
                    WeeksWithDates.Add(weeknum, new List<string>());
                }
                if (!WeeksWithDates[weeknum].Contains(Convert.ToDateTime(failure[0]).ToShortDateString()))
                {
                    WeeksWithDates[weeknum].Add(Convert.ToDateTime(failure[0]).ToShortDateString());
                }
            }

            foreach (int week in WeeksWithDates.Keys)
            {
                HeaderString.AppendFormat("<th colspan='{0}'>Week {1}</th>", (WeeksWithDates[week].Count + 1), week);
                WeeksWithDates[week].Sort();
                foreach (string date in WeeksWithDates[week])
                {
                    SecHeaderString.AppendFormat("<th>{0}</th>", date);
                }
                SecHeaderString.Append("<th>Wk Total</th>");
            }
            HeaderString.Append("<th></th></tr>");
            SecHeaderString.Append("<th>Facility</th></tr>");


            foreach (string facility in facilities) 
            {
                int CountTotal = 0;
                HTMLDataString.AppendFormat("{0}{1}{2}", "<tr><td>",facility,"</td>");
                foreach (int week in WeeksWithDates.Keys)
                {
                    int WeekCount = 0;
                    foreach (string date in WeeksWithDates[week])
                    {
                        int DayCount = 0;
                        foreach(string[] failure in failures){
                            if (failure[3] == facility && Convert.ToDateTime(failure[0]).ToShortDateString() == date)
                            {
                                DayCount++;
                                WeekCount++;
                                CountTotal++;
                            }
                        }
                        if (DayCount != 0)
                        {
                            HTMLDataString.AppendFormat("<td>{0}</td>", DayCount);
                        } 
                        else
                        {
                            HTMLDataString.Append("<td></td>");
                        }
                        
                    }
                    HTMLDataString.AppendFormat("<td class='total'>{0}</td>", WeekCount);
                }
                HTMLDataString.AppendFormat("<td class='total'>{0}</td></tr>", CountTotal);
            }
            return HeaderString.ToString() + SecHeaderString.ToString() + HTMLDataString.Append("</table><br/>").ToString();
        }

        private string GenerateBySiteCountReport(){
            Console.WriteLine("Generating By Site Count Report...");
            StringBuilder ReportHTML = new StringBuilder("<h3>Total Failures By Facility</h3><table><tr><th>Facility</th><th>Failures</th></tr>");
            
            Dictionary<string, int> count = new Dictionary<string, int>();
            foreach (string[] failure in failures)
            {
                if (count.ContainsKey(failure[3]))
                {
                    count[failure[3]]++;
                }
                else
                {
                    count.Add(failure[3], 1);
                }
            }

            foreach (string facility in count.Keys)
            {
                ReportHTML.AppendFormat("<tr><td>{0}</td><td>{1}</td></tr>", facility, count[facility]);
            }
            return ReportHTML.Append("</table><br />").ToString();
        }

        private string GenerateBySiteReport()
        {
            Console.WriteLine("Generating By Site Detailed Report...");
            StringBuilder ReportHTML = new StringBuilder("<h3>Failure Details By Site</h3><table>");
            Dictionary<string, List<string[]>> notices = new Dictionary<string, List<string[]>>();
            foreach (string[] report in failures)
            {
                //check if facility in dict
                if(!notices.ContainsKey(report[3])){
                    notices.Add(report[3], new List<string[]>());
                }
                notices[report[3]].Add(report);
            }

            foreach (string facility in notices.Keys)
            {
                ReportHTML.AppendFormat("<tr><th colspan='7'><h3>{0}</h3></th></tr>", facility);
                ReportHTML.Append("<tr><th>Date</th><th>Environment</th><th>Pat. Location</th><th>Facility</th><th>Destination Group</th><th>Destination</th><th>Address</th></tr>");

                foreach (string[] reports in notices[facility])
                {
                    ReportHTML.Append("<tr>");
                    for (int index = 0; index < reports.Length - 1; index++)
                    {
                        ReportHTML.AppendFormat("<td>{0}</td>", reports[index]);
                    }
                    ReportHTML.Append("</tr>");
                }
            }

            return ReportHTML.Append("</table><br />").ToString();
        }

        static void Main(string[] args)
        {
            System.Diagnostics.Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            Reporter reporter = new Reporter();
            reporter.LoadEmailsToArray();
            Console.WriteLine("Data Collection Completed in " + stopwatch.ElapsedMilliseconds + "ms.");
            reporter.GenerateMailReport();
        }

        public static void GenReport() {
            System.Diagnostics.Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            Reporter reporter = new Reporter();
            reporter.LoadEmailsToArray();
            Debug.WriteLine("Data Collection Completed in " + stopwatch.ElapsedMilliseconds + "ms.");
            reporter.GenerateMailReport();
        }
    }
}
