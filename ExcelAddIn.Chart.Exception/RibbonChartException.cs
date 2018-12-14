/// <summary>
/// WARNING: ANY USE BY YOU OF THE SAMPLE CODE PROVIDED IN THIS FILE IS AT YOUR OWN RISK. 
/// Microsoft provides this code "as is" without warranty of any kind, either express or implied, 
/// including but not limited to the implied warranties of merchantability and/or fitness for a particular purpose.
/// </summary>

using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn.Chart.Exception
{
    public partial class RibbonChartException
    {
        private void RibbonChartException_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonAddChart_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
                    Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            worksheet.Range["A1", "A4"].Value2 = 1;
            worksheet.Range["B1", "B4"].Value2 = 2;
            worksheet.Range["C1", "C4"].Value2 = 3;
            worksheet.Range["D1", "D4"].Value2 = 4;

            // Create a first chart
            double left = 10.00;
            double top = 10.00;
            double width = 200.00;
            double height = 200.00;

            Microsoft.Office.Tools.Excel.Chart chart = worksheet.Controls.AddChart(left, top, width, height, "test" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
            chart.ChartType = Excel.XlChartType.xlColumnClustered;

            Excel.Range cells = worksheet.Range["A1", "D4"];
            chart.SetSourceData(cells);

            chart.PlotArea.Border.Color = ColorTranslator.ToOle(Color.Green);

            try
            {
                chart.PlotArea.Left = 15.00;
                chart.PlotArea.Top = 15.00;
                chart.PlotArea.Width = 150.00;
                chart.PlotArea.Height = 150.00;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
