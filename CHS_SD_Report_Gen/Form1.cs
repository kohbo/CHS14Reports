using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using SLA = SLAReporter;
using PK = PKReporter;

namespace CHSReportGen
{
	public partial class Form1 : Form
	{
		delegate void AppendDelegate(String text);
        delegate string DelegateGetReport();

		String path = "C:\\Users\\Juan.Menendez\\Desktop\\Reports\\";
		Excel.Application oXL;
		Excel.Workbook oWB;
		Excel.Worksheet oWS;
		Excel.Range oRng;
        AppendDelegate ad;
        DelegateGetReport SelectionDel;
		DateTime today;

		public Form1()
		{
			InitializeComponent();
			today = DateTime.Now.AddDays(-1.0);
		}

		private void GenSLA()
		{

            
                try
                {
                    ad = new AppendDelegate(AddToLog);

                    AddToLog("-----   Generating SLA Report   -----");
                    AddToLog("-----      Extracting Data      -----");

                    //Open Excel Application
                    oXL = new Excel.Application
                    {
                        DisplayAlerts = false,
                        Visible = true
                    };
                    oWB = oXL.Workbooks.Add();
                    SLAReporter.Reporter rep = new SLAReporter.Reporter((int)Weeks.Value, endDate.Value, oWB);
                    rep.GetIncidentData();
                    AddToLog("-----    Extraction Complete    -----");
                    AddToLog("-----     Exporting Data to Excel     -----");
                    AddToLog("This may take a while. Please wait...");

                    
                    //TODO Delete Default worksheets

                    //OOP Calls
                    rep.AddOpenedGroupsByWeek();
                    rep.AddGroupByPriority();
//                    rep.AddOpenClosedVolume();
                    rep.AddOpenClosedVolume(priority: "Low", weeks: 99);
                    rep.AddOpenClosedVolume(priority: "Medium", weeks: 99);
                    rep.AddOpenClosedVolume(priority: "High", weeks: 99);
                    rep.AddOpenClosedVolume(priority: "Critical", weeks: 99);
                    rep.AddOpenClosedVolume(priority: "Low");
                    rep.AddOpenClosedVolume(priority: "Medium");
                    rep.AddOpenClosedVolume(priority: "High");
                    rep.AddOpenClosedVolume(priority: "Critical");
                    rep.AddSLAMetByDay();
                    rep.AddOpenClosedVolumeByGroup();

                    AddToLog("-----    SLA Report Complete    -----");
                }
                catch (Exception exc)
                {
                    Console.Write("Exception: " + exc.Message + exc.StackTrace);
                    AddToLog("*****   Error Creating Report   *****");
                    AddToLog("Exception: " + exc.Message + "\n" + exc.StackTrace);
                    oXL.Quit();
                }
                finally
                {
                    //Ensure Excel process is released
                    if (oWS != null) { Marshal.ReleaseComObject(oWS); }
                    if (oWB != null) { Marshal.ReleaseComObject(oWB); }
                    if (oXL != null)
                    {
                        oXL.DisplayAlerts = true;
                        Marshal.ReleaseComObject(oXL);
                    }
                    oWS = null;
                    oWB = null;
                    oXL = null;
                    GC.Collect();
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

		private Excel.Workbook OpenWorkbook(String BookName)
		{
			if (File.Exists(BookName))
			{
                Excel.Workbook WB = oXL.Workbooks.Open(BookName);
                if (!WB.ReadOnly)
                {
                    return WB;
                }
                else
                {
                    AddToLog("Workbook is read-only. Exiting");
                    return null;
                }
			}
			AddToLog("File Not Found: " + BookName);
			return null;
		}

		private void btn_run_Click(object sender, EventArgs e)
		{
            //btn_run.Enabled = false;
			if (sel_report.SelectedItem.ToString().Equals("SLA"))
			{
				Thread runSLA = new Thread(new ThreadStart(GenSLA));
				runSLA.Start();
			}
            if (sel_report.SelectedItem.ToString().Equals("RCA") || sel_report.SelectedItem.ToString().Equals("Huddle"))
            {
                Thread runRCA = new Thread(new ThreadStart(GenRCA));
                runRCA.Start();
            }
            if (sel_report.SelectedItem.ToString().Equals("PK"))
            {
                AddToLog("Running PK Reporter...");
                Thread runPK = new Thread(new ThreadStart(PKReporter.Reporter.GenReport));
                runPK.Start();
            }
            //btn_run.Enabled = true;
		}

        private void SelectedReportChanged(object sender, EventArgs e)
        {
            endDate.Enabled = Weeks.Enabled = (sel_report.SelectedItem == "SLA");
        }

		private void AddToLog(String text)
		{
			if (this.log.InvokeRequired)
			{
				this.Invoke(ad, new object[] { text+"\n" });
			}
			else
			{
                try
                {
                    log.AppendText(text + "\n");
                }
                catch (ObjectDisposedException exc)
                {
                    Console.Write("Exception: " + exc.Message + exc.StackTrace);
                }
			}
					   
		}

        private void GenRCA()
        {
            /*
             * Known Issues:
             * 
             * Includes last non-critical in report
             * Removed Assigned Group cat
             * Copying to other sheet multiple times
            */
            try
            {
                String path = this.path + "RCA\\";
                ad = new AppendDelegate(AddToLog);

                AddToLog("-----   Generating RCA Report   -----");

                //Start new Excel application
                oXL = new Excel.Application();
                oXL.Visible = true;
                if (chk_supwarning.Checked) { oXL.DisplayAlerts = false; }

                SelectionDel = new DelegateGetReport(GetReport);
                //if (GetReport().Equals("RCA"))
                    RCAFormatWorkbook(path + "Open Incident Details.xls");
            }
            catch (Exception exc)
            {
                Console.Write("Exception: " + exc.Message + exc.StackTrace);
                AddToLog("*****   Error Creating Report   *****");
                oXL.Quit();
            }
            finally
            {
                oXL.DisplayAlerts = true;

                //Ensure Excel process is released
                if (oWS != null) { Marshal.ReleaseComObject(oWS); }
                if (oWB != null) { Marshal.ReleaseComObject(oWB); }
                if (oXL != null) { Marshal.ReleaseComObject(oXL); }
                oWS = null;
                oWB = null;
                oXL = null;
                GC.Collect();
            }
            AddToLog("-----   Report Complete   -----");
        }

        private void RCAFormatWorkbook(String BookName)
        {
            List<String> ReportColumns = new List<string>(){
                "Incident Number",
                "Submit Time",
                "Submit Date",
                "Priority",
                "Status",
                "Company",
                "Submit Time",
                "Submit Date",
                "Last Resolved Time",
                "Last Modified Date",
                "Summary",
                "Assigned Group"
            };

            oWB = oXL.Workbooks.Open(BookName);
            oWS = oWB.ActiveSheet;

            //Remove Title Rows
            int ColIndex = FindColumnIndex("Priority",oWS.Rows[1]);

            AddToLog("Removing non-criticals...");
            //row to be critical
            int NotesLoc = FindColumnIndex("Notes", oWS.Rows[1]);
            int ResolutionLoc = FindColumnIndex("Resolution", oWS.Rows[1]);
            for (int index = 2; index <= oWS.UsedRange.Rows.Count; index++)
            {
                //rows not critical
                for (int index2 = index; index2 <= oWS.UsedRange.Rows.Count; index2++)
                {
                    //if row is for critical incident
                    if (oWS.Rows[index2].Columns[ColIndex].Value2.ToString() == "Critical")
                    {
                        //if first row is critical, don't delete
                        if (index == index2)
                        {
                            break;
                        }
                        //delete range of non-crtiticals
                        oWS.get_Range("A" + index, "A" + (index2 - 1)).EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        break;
                    }
                    //if reached end of report
                    if (index2 == oWS.UsedRange.Rows.Count)
                    {
                        oWS.get_Range("A" + index, "A" + index2).EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        break;
                    }
                }

                //Find nuance issues and highlight
                if (oWS.UsedRange.Rows[index].Columns[NotesLoc].Value2 != null)
                {
                    if (oWS.UsedRange.Rows[index].Columns[NotesLoc].Value2.Contains("Nuance"))
                    {
                        oWS.UsedRange.Rows[index].Interior.Color = System.Drawing.Color.PaleVioletRed;
                    }
                }
                if (oWS.UsedRange.Rows[index].Columns[ResolutionLoc].Value2 != null)
                {
                    if (oWS.UsedRange.Rows[index].Columns[ResolutionLoc].Value2.Contains("TRNS"))
                    {
                        oWS.UsedRange.Rows[index].Interior.Color = System.Drawing.Color.PaleVioletRed;
                    }
                }
            }
            AddToLog("Done");

            AddToLog("Moving columns to new sheet...");
            oWB.Worksheets.Add();
            Excel._Worksheet oWSNew = oWB.Worksheets[1];

            int PasteIndex = 1;
            bool UnmergeFlag = true;

            //move data to new sheet in correct column order
            //oWS.Columns.ClearFormats();
            foreach(String col in ReportColumns)
            {
                ColIndex = FindColumnIndex(col, oWS.Rows[1]);
                if (ColIndex != 0)
                {
                    if (oWS.UsedRange.Columns[ColIndex].Cells[1].Value2.ToString() == "Company" && UnmergeFlag)
                    {
                        //oWS.UsedRange.Columns.Mer
                        UnmergeFlag = false;
                        oWS.UsedRange.Columns[ColIndex].UnMerge();
                        oWS.UsedRange.Columns[ColIndex + 1].EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    }
                    oWS.UsedRange.Columns[ColIndex].Copy(oWSNew.Columns[PasteIndex++]);
                }
            }
            AddToLog("Done");

            //formatting corrections
            oWSNew.UsedRange.Columns.ColumnWidth = 22.57;
            //oWSNew.Columns[2].NumberFormat("m/d/yy");
           // oWSNew.Columns[6].NumberFormat("hh:mm AM/PM");
           // oWSNew.Columns[7].NumberFormat("hh:mm AM/PM");
            oWSNew.Columns.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oWSNew.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        private int FindColumnIndex(String ColName, Excel.Range Row)
        {
            for (int index = 1; index <= Row.Columns.Count; index++)
            {
                if (Row.Columns[index].Value2 == ColName)
                {
                    return index;
                }
            }
            return 0;
        }

        private String GetReport()
        {
            if (this.sel_report.InvokeRequired)
            {
                return SelectionDel.Invoke();
            }
            else
            {
                return sel_report.SelectedItem.ToString();
            }
        }
	}
}
