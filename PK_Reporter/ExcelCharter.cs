using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PKReporter
{
    class ExcelCharter
    {
        List<string[]> failures;
        Excel.Application oExcelApp;
        Excel.Workbook oWB;
        Excel.Worksheet oWS;
        Excel.Range range;
        Excel.Chart chart;

        public ExcelCharter(List<string[]> failures)
        {
            this.failures = failures;
            oExcelApp = new Excel.Application();
            oWB = oExcelApp.Workbooks.Add();
            oWS = oWB.ActiveSheet;
            oExcelApp.Visible = true;
            this.CopyDataToSheet();
            this.CreateChart();
        }

        ~ExcelCharter()
        {
            Marshal.ReleaseComObject(oExcelApp);
            Marshal.ReleaseComObject(oWB);
            Marshal.ReleaseComObject(oWS);
            oExcelApp = null;
            oWB = null;
            oWS = null;
        }     

        private void CopyDataToSheet()
        {
            int col = 0;
            int row = 0;
            Dictionary<string, Dictionary<string, int>> totals = new Dictionary<string, Dictionary<string, int>>();
            List<string> types =  new List<string>();
            foreach (string[] failure in failures)
            {
                //add facilities to keys
                if (!totals.ContainsKey(failure[3]))
                {
                    totals.Add(failure[3], new Dictionary<string, int>());
                }
                //add failure types to types
                if (!types.Contains(failure[4]))
                {
                    types.Add(failure[4]);
                }
                //add failure type to facility
                if (!totals[failure[3]].ContainsKey(failure[4]))
                {
                    totals[failure[3]].Add(failure[4], 1);
                }
                else
                {
                    totals[failure[3]][failure[4]]++;
                }
            }

            range = oWS.Rows[1];
            col = 2;
            foreach (string type in types)
            {
                range.Columns[col++].Value2 = type;
            }

            range = oWS.Columns[1];
            row = 2;
            col = 2;
            foreach (string facility in totals.Keys)
            {
                col = 2;
                range.Rows[row].Value2 = facility;
                foreach (string type in types)
                {
                    if (totals[facility].ContainsKey(type))
                    {
                        range.Columns[col].Rows[row].Value2 = totals[facility][type];
                    }
                    else
                    {
                        range.Columns[col].Rows[row].Value2 = 0;
                    }
                    col++;
                }
                row++;
            }
        }

        private void CreateChart()
        {
            chart = oWS.ChartObjects().Add(0,0,800,500).Chart;
            chart.SetSourceData(oWS.UsedRange);
            chart.ChartType = Excel.XlChartType.xlColumnClustered;
            chart.ChartStyle = 17;
            foreach (Excel.ChartGroup group in chart.ChartGroups())
            {
                group.GapWidth = 0;
                group.Overlap = 0;
            }
        }

        public Excel.Chart getChart()
        {
            return chart;
        }
    }
}
