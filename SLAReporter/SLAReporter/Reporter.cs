using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;

namespace SLAReporter
{
    public class Reporter
    {
        public List<Incident> incidents;
        public List<string> Groups;
        DateTime[] ByGroupWeeks;
        DateTime endDate;
        int Weeks;
        readonly Excel.Workbook oWB;
        private SortedDictionary<DateTime, int> Days_Opened;
        private SortedDictionary<DateTime, int> Days_Resolved;

        public Reporter(int Weeks, DateTime endDate, Excel.Workbook oWB)
        {
            this.oWB = oWB;
            this.Weeks = Weeks;
            this.endDate = endDate;
            ByGroupWeeks = new DateTime[Weeks + 1];
            for (int i = ByGroupWeeks.Length - 1; i >= 0; i--)
            {
                ByGroupWeeks[i] = (i == ByGroupWeeks.Length - 1) ? endDate.AddDays(1) : endDate.AddDays(-7 * (ByGroupWeeks.Length - 1 - i));
            }
        }

        public void GetIncidentData()
        {
            incidents = new List<Incident>();
            Groups = new List<string>();
            Days_Opened = new SortedDictionary<DateTime, int>();
            Days_Resolved = new SortedDictionary<DateTime, int>();

            using (TextFieldParser parser = new TextFieldParser(@".\RemedyReport\Incidents.csv"))
            {
                string[] top_row = {};

                parser.SetDelimiters(",");
                if (!parser.EndOfData)
                {
                    top_row = parser.ReadFields();
                }

                Dictionary<string, int> headers = new Dictionary<string, int> {
                        {"number" , GetColumnIndex("Incident_Number", top_row)},
                        {"facility", GetColumnIndex("Company", top_row)},
                        {"group", GetColumnIndex("Assigned_Group", top_row)},
                        {"product", GetColumnIndex("Product_Name", top_row)},
                        {"summary", GetColumnIndex("Summary", top_row)},
                        {"priority", GetColumnIndex("Priority", top_row)},
                        {"submitted", GetColumnIndex("Submit_Time", top_row)},
                        {"status", GetColumnIndex("Status", top_row)},
                        {"resolved", GetColumnIndex("Last_Resolved_Time", top_row)},
                        {"days_open", GetColumnIndex("Open_Time_Days", top_row)},
                      };

                List<string> GroupExclusions = new List<string>
                {
                        "CHS14 Implementation",
                        "CHS14 Interface",
                        "CHS14 MEDHOST EDIS",
                        "CHS14 NextGen Finance",
                        "CHS14 NextGen Support",
                        "CHS14 PatientKeeper - MedRec",
                        "CHS14 PMM - McKesson Supply Chain Management-MSCM",
                        "CHS14 Project Managers",
                        "CHS14 Release Management",
                        "CHS14 Sarasota Shared Services Center",
                };

                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    //TODO add dictionary to filter products
                    if (fields[headers["product"]] != "NimBUS" && !GroupExclusions.Contains(fields[headers["group"]]))
                    {
                        Incident inc = new Incident(
                                            fields[headers["number"]],
                                            fields[headers["facility"]],
                                            fields[headers["group"]],
                                            fields[headers["product"]],
                                            fields[headers["summary"]],
                                            fields[headers["priority"]],
                                            Convert.ToDateTime(fields[headers["submitted"]]),
                                            fields[headers["status"]],
                                            (fields[headers["resolved"]] == "") ? DateTime.MinValue : Convert.ToDateTime(fields[headers["resolved"]]), //resolved
                                            Convert.ToDouble(fields[headers["days_open"]]) //days open
                                        );
                        if (!Groups.Contains(fields[headers["group"]]))
                        {
                            Groups.Add(fields[headers["group"]]);
                        }
                        incidents.Add(inc);

                        //Incident Opened Count By Day
                        if (!Days_Opened.ContainsKey(inc.submitted.Date))
                        {
                            Days_Opened.Add(inc.submitted.Date, 1);
                        }
                        else
                        {
                            Days_Opened[inc.submitted.Date]++;
                        }

                        //Incident Resolved Count By Day
                        if (inc.last_resolved != DateTime.MinValue)
                        {
                            if (!Days_Resolved.ContainsKey(inc.last_resolved.Date))
                            {
                                Days_Resolved.Add(inc.last_resolved.Date, 1);
                            }
                            else
                            {
                                Days_Resolved[inc.last_resolved.Date]++;
                            }
                        }
                    }
                }
                Groups.Sort();
            }
        }

        private int GetColumnIndex(string SearchField, string[] row)
        {
            for (int index = 0; index < row.Length; index++)
            {
                if (row[index] == SearchField)
                {
                    return index;
                }
            }
            return -1;
        }

        public void AddOpenedGroupsByWeek()
        {
            Excel.Worksheet oWS = oWB.Worksheets.Add();
            oWS.Name = "Criticals By Group By Week";

            SortedDictionary<string, int[]> OpenedByGroupByWeek = new SortedDictionary<string, int[]>();

            foreach (Incident inc in incidents)
            {
                //extract OpenedByGroupByWeek
                if (!OpenedByGroupByWeek.ContainsKey(inc.group)) { OpenedByGroupByWeek.Add(inc.group, PopulateArray<int>(new int[Weeks], 0)); }
                for (int index = 0; index < ByGroupWeeks.Length - 1; index++)
                {
                    if (inc.submitted >= ByGroupWeeks[index] && inc.submitted < ByGroupWeeks[index + 1] && inc.priority == "Critical")
                    {
                        OpenedByGroupByWeek[inc.group][index]++;
                    }
                }
            }

            //By Group By Week -> Excel
            oWS.Range["A1"].Value2 = "Groups";
            int col = 2;
            int row = 1;
            foreach (DateTime week in ByGroupWeeks)
            {
                if (week == ByGroupWeeks.Last()) { break; }
                oWS.Rows[row].Columns[col].Value2 = week.ToShortDateString() + " - " + week.AddDays(6).ToShortDateString();
                oWS.Columns[col++].ColumnWidth = 19.5;
                //oWS.Rows[row].Columns[col++].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            }

            row = 2;
            foreach (string group in OpenedByGroupByWeek.Keys)
            {
                if (DetermineNoTickets(OpenedByGroupByWeek[group])) continue;
                col = 1;
                oWS.Rows[row].Columns[col].Value2 = group;
                foreach (int count in OpenedByGroupByWeek[group])
                {
                    oWS.Rows[row].Columns[++col].Value2 = Convert.ToString(count);
                }
                row++;
            }
            oWS.Columns[1].AutoFit();
            oWS.Columns[1].Rows[row + 2].Value2 = "Problem Identification and Management";

            //TODO Separate these methods
            AddTopTenCrits(oWB, OpenedByGroupByWeek);
        }

        public void AddTopTenCrits(Excel.Workbook oWB, SortedDictionary<string, int[]> OpenedByGroupByWeek)
        {
            Dictionary<string, int> TopTenCriticalGroups = new Dictionary<string, int>();

            Excel.Worksheet oWS = oWB.Worksheets.Add();

            
            //ExtractTopTenCriticals
            for (int i = 0; i < 10; i++)
            {
                string biggest_s = "";
                int biggest_c = 0;
                foreach (string group in OpenedByGroupByWeek.Keys)
                {
                    if (OpenedByGroupByWeek[group][Weeks - 1] >= biggest_c && !TopTenCriticalGroups.ContainsKey(group))
                    {
                        biggest_s = group;
                        biggest_c = OpenedByGroupByWeek[group][Weeks - 1];
                    }
                }
                TopTenCriticalGroups.Add(biggest_s, biggest_c);
            }

            //Top Ten Groups -> Excel
            oWS.Name = "Top Ten Group Criticals";
            oWS.Range["A1"].Value2 = "Group";
            oWS.Range["B1"].Value2 = "Criticals Last Week";
            int row = 2;
            foreach (string group in TopTenCriticalGroups.Keys)
            {
                oWS.Rows[row].Columns[1].Value2 = group;
                oWS.Rows[row].Columns[2].Value2 = TopTenCriticalGroups[group].ToString();
                row++;
            }
        }

        public void AddGroupByPriority()
        {
            SortedDictionary<string, int[]> ByGroupByPriority = new SortedDictionary<string, int[]>();
            Dictionary<string, int> PriorityShifter = new Dictionary<string, int>
                    {
                        {"Low", 0},
                        {"Medium", 3},
                        {"High", 6},
                        {"Critical", 9}
                    };
            foreach (Incident inc in incidents)
            {
                //extract ByGroupByPriority
                if (!ByGroupByPriority.ContainsKey(inc.group)) { ByGroupByPriority.Add(inc.group, new int[12] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }); }
                //add existing
                if (inc.submitted <= endDate && inc.last_resolved == DateTime.MinValue)
                {
                    ByGroupByPriority[inc.group][PriorityShifter[inc.priority]]++;
                }
                //add closed
                if (inc.last_resolved >= endDate.AddDays(-7) && inc.last_resolved <= endDate)
                {
                    ByGroupByPriority[inc.group][PriorityShifter[inc.priority] + 1]++;
                }
                //add opened
                if (inc.submitted >= endDate.AddDays(-7) && inc.submitted <= endDate)
                {
                    ByGroupByPriority[inc.group][PriorityShifter[inc.priority] + 2]++;
                }
            }

            //By Group By Priority -> Excel
            Excel.Worksheet oWS = oWB.Worksheets.Add();
            Excel.Range oRng;
            oWS.Name = "Group Criticals By Priority";
            oWS.get_Range("A2").Value2 = "Group";
            int col = 2;
            foreach (string pri in PriorityShifter.Keys)
            {
                //if (pri == PriorityShifter.Keys.Last()) { break; }
                oRng = oWS.Range[oWS.Rows[1].Columns[col], oWS.Rows[1].Columns[col + 2]];
                MergeAndCenter(oRng, pri);
                col += 3;
            }

            for (int i = 2; i < 12; )
            {
                oWS.Rows[2].Columns[i++].Value2 = "Existing";
                oWS.Rows[2].Columns[i++].Value2 = "Closed";
                oWS.Rows[2].Columns[i++].Value2 = "New";
            }

            int row = 3;
            foreach (string group in ByGroupByPriority.Keys)
            {
                if (DetermineNoTickets(ByGroupByPriority[group])) continue;
                oWS.Rows[row].Columns[1].Value2 = group;
                col = 2;
                foreach (int count in ByGroupByPriority[group])
                {
                    oWS.Rows[row].Columns[col++].Value2 = count.ToString();
                }
                row++;
            }
            oWS.Columns[1].AutoFit();
        }

        public void AddOpenClosedVolume(string priority = "all", int weeks = 8)
        {
            Excel.Worksheet oWS = oWB.Worksheets.Add();

            //Open vs Closed by Week with Volume Line
            SortedDictionary<DateTime, int[]> OpenClosedWeeksDelimiters = new SortedDictionary<DateTime, int[]>();
            for (DateTime index = (weeks == 99) ? new DateTime(DateTime.Now.Year, 1, 1) : endDate.Date.AddDays(-7*weeks);
                index < endDate.Date;
                index = index.AddDays(7))
            {
                OpenClosedWeeksDelimiters.Add(index, new int[3] {0, 0, 0});
            }

            foreach (Incident inc in incidents)
            {
                if (inc.priority != priority && priority != "all") continue;
                foreach (DateTime week in OpenClosedWeeksDelimiters.Keys)
                {
                    if (inc.submitted >= week && inc.submitted < week.AddDays(7))
                    {
                        OpenClosedWeeksDelimiters[week][0]++;
                    }
                    if (inc.last_resolved >= week && inc.last_resolved < week.AddDays(7))
                    {
                        OpenClosedWeeksDelimiters[week][1]++;
                    }
                    if (inc.submitted <= week.AddDays(7) &&
                        (inc.last_resolved > week.AddDays(7) || inc.last_resolved == DateTime.MinValue))
                    {
                        OpenClosedWeeksDelimiters[week][2]++;
                    }
                }
            }

            //Open vs Closed by Week with Volume -> Excel
            oWS.Name = "OpenClosed" + ((weeks == 99) ? "Year" : "Last" + weeks) + priority.ToUpper();
            oWS.Range["A1"].Value2 = "Date";
            oWS.Range["B1"].Value2 = "Created";
            oWS.Range["C1"].Value2 = "Resolved";
            oWS.Range["D1"].Value2 = "Open";
            int row = 2;

            foreach (KeyValuePair<DateTime, int[]> week in OpenClosedWeeksDelimiters)
            {
                oWS.Rows[row].Columns[1].Value2 = week.Key.ToShortDateString();
                oWS.Rows[row].Columns[2].Value2 = week.Value[0];
                oWS.Rows[row].Columns[3].Value2 = week.Value[1];
                oWS.Rows[row].Columns[4].Value2 = week.Value[2];
                row++;
            }

            Excel.ChartObjects cObj = oWS.ChartObjects();
            Excel.Chart chart = cObj.Add(100, 10, 500, 400).Chart;
            chart.SetSourceData(oWS.UsedRange);
            chart.ChartType = Excel.XlChartType.xlColumnClustered;
            chart.HasTitle = true;
            chart.ChartTitle.Text = priority[0].ToString().ToUpper() + priority.Substring(1);
            Excel.Axis cAxis = chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            cAxis.CategoryType = Excel.XlCategoryType.xlCategoryScale;
            Excel.Series oSC = chart.SeriesCollection("Open");
            oSC.ChartType = Excel.XlChartType.xlLine;
            oSC.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbBlack;
            chart.HasAxis[Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary] = true;
            oSC.AxisGroup = Excel.XlAxisGroup.xlSecondary;
        }

        public void AddOpenClosedVolumeByGroup(int weeks = 8)
        {
            Excel.Worksheet oWS = oWB.Worksheets.Add();
            oWS.Name = "OpenClosed By Group";
            int col = 2;
            int indexc = 0;

            SortedDictionary<DateTime, int[]> OpenClosedWeeksDelimiters = new SortedDictionary<DateTime, int[]>();
            for (DateTime index = ((weeks == 99) ? new DateTime(DateTime.Now.Year, 1, 1) : endDate.Date.AddDays(-7 * weeks)); index < endDate.Date; index = index.AddDays(7))
            {
                OpenClosedWeeksDelimiters.Add(index, PopulateArray(new int[12], 0));
            }

            oWS.Range["A3"].Value2 = "Week";
            oWS.Range["A4", "A" + (OpenClosedWeeksDelimiters.Keys.Count+3)].Value2 =
                OneToTwoD(OpenClosedWeeksDelimiters.Keys.ToArray());
            oWS.Range["A4", "A" + (OpenClosedWeeksDelimiters.Keys.Count+3)].NumberFormat = "m/d/yyyy";

            Dictionary<string, int> PriorityShift = new Dictionary<string, int>
            {
                {"Critical", 0},
                {"High", 3},
                {"Medium", 6},
                {"Low", 9}
            };

            foreach (string group in Groups)
            {
                int row = 4;
                foreach (Incident inc in incidents)
                {
                    if (inc.group != group) continue;
                    foreach (DateTime week in OpenClosedWeeksDelimiters.Keys)
                    {
                        //add opened
                        if (inc.submitted >= week && inc.submitted < week.AddDays(7))
                        {
                            (OpenClosedWeeksDelimiters[week][PriorityShift[inc.priority]])++;
                        }
                        //add closed
                        if (inc.last_resolved >= week && inc.last_resolved < week.AddDays(7))
                        {
                            (OpenClosedWeeksDelimiters[week][PriorityShift[inc.priority] + 1])++;
                        }
                        //add open
                        if (inc.submitted <= week.AddDays(7) &&
                            (inc.last_resolved > week.AddDays(7) || inc.last_resolved == DateTime.MinValue))
                        {
                            (OpenClosedWeeksDelimiters[week][PriorityShift[inc.priority] + 2])++;
                        }
                    }
                }

                foreach (var date in OpenClosedWeeksDelimiters.Keys)
                {
                    for (int x = 0; x < 12; x++)
                    {
                        oWS.Cells[row, col + x].Value2 = OpenClosedWeeksDelimiters[date][x].ToString();
                    }
                    //oWS.Range[oWS.Cells[row, col], oWS.Cells[row, col + 11]].Value2 =
                    //    OneToTwoD(OpenClosedWeeksDelimiters[date].ToArray());
                    row++;
                }

                oWS.Cells[1, col].Value2 = group;
                foreach (var priority in PriorityShift.Keys)
                {
                    oWS.Cells[2, col].Value2 = priority;
                    oWS.Cells[3, col++].Value2 = "Created";
                    oWS.Cells[3, col++].Value2 = "Resolved";
                    oWS.Cells[3, col++].Value2 = "Open";
                }

                OpenClosedWeeksDelimiters = new SortedDictionary<DateTime, int[]>();
                for (DateTime index = ((weeks == 99) ? new DateTime(DateTime.Now.Year, 1, 1) : endDate.Date.AddDays(-7 * weeks)); index < endDate.Date; index = index.AddDays(7))
                {
                    OpenClosedWeeksDelimiters.Add(index, PopulateArray(new int[12], 0));
                }

                AddChart(oWS, 10, 10 + (0 + indexc * 900), 600, 400, group + ' ' + "Critical", oWS.Range[oWS.Cells[3, col - 12], oWS.Cells[3 + OpenClosedWeeksDelimiters.Count, col - 10]]);
                AddChart(oWS, 650, 10 + (0 + indexc * 900), 600, 400, group + ' ' + "High", oWS.Range[oWS.Cells[3, col - 9], oWS.Cells[3 + OpenClosedWeeksDelimiters.Count, col - 7]]);
                AddChart(oWS, 10, 10 + (450 + indexc * 900), 600, 400, group + ' ' + "Medium", oWS.Range[oWS.Cells[3, col - 6], oWS.Cells[3 + OpenClosedWeeksDelimiters.Count, col - 4]]);
                AddChart(oWS, 650, 10 + (450 + indexc * 900), 600, 400, group + ' ' + "Low", oWS.Range[oWS.Cells[3, col - 3], oWS.Cells[3 + OpenClosedWeeksDelimiters.Count, col - 1]]);

                indexc++;
            }
        }

        public void AddSLAMetByDay(string priority = "all")
        {

            Excel.Worksheet oWS = oWB.Worksheets.Add();
            oWS.Name = "SLA Met by Priority";

            SortedDictionary<string, Dictionary<DateTime, int[]>> SLAMetByPriorityByDay = new SortedDictionary<string, Dictionary<DateTime, int[]>>
                    {
                        {"Low", new Dictionary<DateTime, int[]>()},
                        {"Medium", new Dictionary<DateTime, int[]>()},
                        {"High", new Dictionary<DateTime, int[]>()},
                        {"Critical", new Dictionary<DateTime, int[]>()}
                    };
            Dictionary<string, int> SLAStandards = new Dictionary<string, int> {
                        {"Low", 240},
                        {"Medium", 168},
                        {"High", 48},
                        {"Critical", 8}
                    };

            foreach (Incident inc in incidents)
            {
                //SLAMetByPriorityByDay
                if (!SLAMetByPriorityByDay[inc.priority].ContainsKey(inc.last_resolved.Date))
                {
                    SLAMetByPriorityByDay[inc.priority].Add(inc.last_resolved.Date, new int[] { 0, 1 });
                }
                else
                {
                    SLAMetByPriorityByDay[inc.priority][inc.last_resolved.Date][1]++;
                }
                if ((inc.last_resolved - inc.submitted).TotalHours <= SLAStandards[inc.priority] && (inc.last_resolved - inc.submitted).Hours > 0)
                {
                    SLAMetByPriorityByDay[inc.priority][inc.last_resolved.Date][0]++;
                }
            }

            //SLA Met By Priority By Day -> Excel
            oWS.Range["A1"].Value2 = "Date";
            int col = 2;
            foreach (string pri in SLAMetByPriorityByDay.Keys)
            {
                oWS.Rows[1].Columns[col++].Value2 = pri;
            }
            int row = 2;
            foreach (DateTime day in Days_Resolved.Keys)
            {
                oWS.Rows[row].Columns[1].Value2 = day.ToShortDateString();
                col = 2;
                foreach (string pri in SLAMetByPriorityByDay.Keys)
                {
                    if (SLAMetByPriorityByDay[pri].ContainsKey(day))
                    {
                        oWS.Rows[row].Columns[col+4].Value2 = ((float)SLAMetByPriorityByDay[pri][day][0] / (float)SLAMetByPriorityByDay[pri][day][1]).ToString();
                        oWS.Rows[row].Columns[col++].Value2 = ((float)SLAMetByPriorityByDay[pri][day][0] / (float)Days_Resolved[day]).ToString();
                    }
                    else
                    {
                        oWS.Rows[row].Columns[col+4].Value2 = "0";
                        oWS.Rows[row].Columns[col++].Value2 = "0";
                    }
                }
                row++;
            }
        }

        private T[] PopulateArray<T>(T[] arr, T value)
        {
            for (int i = 0; i < arr.Length; i++)
            {
                arr[i] = value;
            }
            return arr;
        }
        
        private void MergeAndCenter(Excel.Range tRng, string value)
        {
            tRng.Merge();
            tRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            tRng.Value2 = value;
        }

        private T[,] OneToTwoD<T>(T[] arr)
        {
            T[,] TwoDArr = new T[arr.Length, 1];
            for (int i = 0; i < arr.Length; i++)
            {
                TwoDArr[i,0] = arr[i];
            }
            return TwoDArr;
        }

        public static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        private bool DetermineNoTickets(int[] totals)
        {
            foreach (int i in totals)
            {
                if (i > 0) return false;
            }

            return true;
        }

        private void AddChart(Excel.Worksheet oWS, double left, double top, double width, double height, string title, Excel.Range src)
        {
            Excel.ChartObjects oChartObjs = oWS.ChartObjects();
            Excel.Chart oChart = oChartObjs.Add(left, top, width, height).Chart;
            oChart.HasTitle = true;
            oChart.ChartTitle.Text = title;
            oChart.SetSourceData(src);
            Excel.Axis cAxis = oChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            cAxis.CategoryType = Excel.XlCategoryType.xlCategoryScale;
            cAxis.CategoryNames = oWS.Range["A4", "A56"];
            Excel.Series oSC = oChart.SeriesCollection("Open");
            oSC.ChartType = Excel.XlChartType.xlLine;
            oSC.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbBlack;
            oChart.HasAxis[Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary] = true;
            oSC.AxisGroup = Excel.XlAxisGroup.xlSecondary;
        }
    }
}
