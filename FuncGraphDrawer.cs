using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace archlab9
{
    internal class FuncGraphDrawer
    {
        internal void DrawGraph(int start, int end)
        {
            var excel = new Excel.Application();
            var book = excel.Workbooks.Add();
            var sheet = book.Worksheets.Add();


            for (int i = 0; i < end - start; i++)
            {
                sheet.Cells[i + 1, 1] = start + i;
                sheet.Cells[i + 1, 2].Formula = $"=A{i + 1} * A{i + 1}";
            }

            var chartObjs = (ChartObjects)sheet.ChartObjects();
            var chartObj = chartObjs.Add(5, 50, 300, 300);
            var chart = chartObj.Chart;
            var xS = sheet.Range[$"A1:A{end - start}"];
            var yS = sheet.Range[$"B1:B{end - start}"];
            chart.ChartType = Excel.XlChartType.xlXYScatterSmooth;

            var seriesCol = (SeriesCollection)chart.SeriesCollection(Type.Missing);

            var series = seriesCol.NewSeries();

            series.XValues = xS;
            series.Values = yS;


            void TrySave()
            {
                try
                {
                    book.SaveAs("C:\\Users\\Андрей Лузгин\\OneDrive\\Desktop\\учеба\\архитектура ИС\\9\\graph");
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }

            TrySave();
            book.Close();

        }
    }
}
