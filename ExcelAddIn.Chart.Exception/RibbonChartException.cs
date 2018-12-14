using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Drawing;
using System.Windows.Forms;

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

            Microsoft.Office.Tools.Excel.Chart chart = worksheet.Controls.AddChart(left, top, width, height, "test" + DateTime.Now.ToString("yyyymmddhhssfff"));
            chart.ChartType = Excel.XlChartType.xl3DPie;

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
