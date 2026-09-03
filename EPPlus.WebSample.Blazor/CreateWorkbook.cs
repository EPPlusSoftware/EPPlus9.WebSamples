using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Table;
using System.Dynamic;
using System.Globalization;
using System.Net.Http.Json;
using System.Reflection;
using System.Text.Json;

namespace EPPlus.WebSample.Blazor
{
    /// <summary>
    /// This class creates the workbook used in the Blazor sample application.
    /// </summary>
    public static class CreateWorkbook
    {
        public static async Task<ExcelPackage> CreateSampleWorkbook(IEnumerable<ExpandoObject>? cities, eChartType chartType, ePresetChartStyle chartStyle)
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPlus Blazor Sample");
            var package = new ExcelPackage();
            // Add a named style for hyperlinks
            var ns = package.Workbook.Styles.CreateNamedStyle("Hyperlink");
            ns.BuildInId = 8; //Build in type 8 is Hyperlink
            ns.Style.Font.Color.SetColor(eThemeSchemeColor.Hyperlink);
            ns.Style.Font.UnderLine = true;

            var sheet = package.Workbook.Worksheets.Add("Cities");

            await LoadCitiesListFromJavascript(sheet, cities);

            // add a hyperlink to next worksheet
            var nextWorksheetHyperlink = new ExcelHyperLink("'Fx Rates'!A1", "Go to next worksheet");
            sheet.Cells["A13"].Hyperlink = nextWorksheetHyperlink;
            sheet.Cells["A13"].StyleName = "Hyperlink";

            // Adjust the column width after the widest text in the cells
            sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column].AutoFitColumns();

            // read png image from embedded resource
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("EPPlus.WebSample.Blazor.resources.EPPlus-logo-full.png");
            var pic = sheet.Drawings.AddPicture("epplogo", stream);
            pic.SetSize(50);
            pic.SetPosition(0, 10, 0, 10);


            // call an api endpoint to retrieve json data with fx rates.
            var client = new HttpClient();
            var url = new Uri("https://localhost:44354/api/fxrates");
            var response = await client.GetAsync(url);
            var fxRatesJson = await response.Content.ReadFromJsonAsync<string>();
            var fxRates = JsonSerializer.Deserialize<IEnumerable<ExpandoObject>>(fxRatesJson);
            CreateTableAndChartWorksheet(package, fxRates, chartType, chartStyle);

            // finally serialize the workbook to a byte array and use the BlazorDownloadFile function to create a link for download.
            return package;
        }    
        private static async Task LoadCitiesListFromJavascript(ExcelWorksheet sheet, IEnumerable<ExpandoObject>? cities)
        {
            // get some json data from a javascript function and convert it into ExpandoObjects (dynamic objects)
            // load the dynamic objects into EPPlus using the LoadFromDictionaries method.
            var loadedRange = sheet.Cells["A5"].LoadFromDictionaries(cities, x =>
            {
                x.TableStyle = OfficeOpenXml.Table.TableStyles.Medium13;
                x.PrintHeaders = true;
                x.DataTypes = new eDataTypes[] { eDataTypes.String, eDataTypes.Number, eDataTypes.String, eDataTypes.Number };
            });
            var table = sheet.Tables.GetFromRange(loadedRange);
            table.ShowTotal = true;
            table.Columns[0].TotalsRowLabel = "Average";
            table.Columns[1].TotalsRowFunction = RowFunctions.Average;
            table.Columns[3].TotalsRowFunction = RowFunctions.Average;
            sheet.Cells[table.Range.Start.Row, 2, table.Range.End.Row, 2].Style.Numberformat.Format = "#,##0";
            sheet.Cells[table.Range.Start.Row, 4, table.Range.End.Row, 4].Style.Numberformat.Format = "#,##0 \"km2\"";
            // Calculate the formulas in the worksheet, in this case the average functions of the table's total row
            sheet.Calculate();
        }

        private static void CreateTableAndChartWorksheet(ExcelPackage p, IEnumerable<ExpandoObject>? data, eChartType chartType, ePresetChartStyle chartStyle)
        {
            var ws = p.Workbook.Worksheets.Add("FX Rates");
            ws.View.ShowGridLines = false;

            var range = ws.Cells["A20"].LoadFromDictionaries(data, x =>
            {
                x.TableStyle = null;
                x.DataTypes = new[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number };
                x.PrintHeaders = true;
                x.Culture = CultureInfo.InvariantCulture;
            });

            ws.Cells["A:A"].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells["B:F"].Style.Numberformat.Format = "#,##0.00";

            ws.Cells[22, 7, range.End.Row, 11].Formula = "B$21/B22-1";
            ws.Cells["G:K"].Style.Numberformat.Format = "0.00%";

            ws.Cells["B20"].Value = "USD/SEK";
            ws.Cells["C20"].Value = "USD/EUR";
            ws.Cells["D20"].Value = "USD/INR";
            ws.Cells["E20"].Value = "USD/CNY";
            ws.Cells["F20"].Value = "USD/DKK";

            ws.Cells["G20"].Value = "USD/SEK %";
            ws.Cells["H20"].Value = "USD/EUR %";
            ws.Cells["I20"].Value = "USD/INR %";
            ws.Cells["J20"].Value = "USD/CNY %";
            ws.Cells["K20"].Value = "USD/DKK %";

            //Add a table over the range including the .
            var tbl = ws.Tables.Add(ws.Cells[20, 1, range.End.Row, 11], "Table2");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Dark6;

            ws.View.FreezePanes(21, 1);

            CreateChart(ws, range, chartType, chartStyle);

            ws.Calculate();
            ws.Cells.AutoFitColumns();
        }

        private static void CreateChart(ExcelWorksheet ws, ExcelRangeBase range, eChartType chartType, ePresetChartStyle chartStyle)
        {
            var chart = ws.Drawings.AddChart("Chart1", chartType);
            var s1 = chart.Series.Add(ws.Cells[21, 7, range.End.Row, 7], ws.Cells[21, 1, range.End.Row, 1]);
            s1.HeaderAddress = ws.Cells["G20"];

            var s2 = chart.Series.Add(ws.Cells[21, 8, range.End.Row, 8], ws.Cells[21, 1, range.End.Row, 1]);
            s2.HeaderAddress = ws.Cells["H20"];

            var s3 = chart.Series.Add(ws.Cells[21, 9, range.End.Row, 9], ws.Cells[21, 1, range.End.Row, 1]);
            s3.HeaderAddress = ws.Cells["I20"];

            var s4 = chart.Series.Add(ws.Cells[21, 10, range.End.Row, 10], ws.Cells[21, 1, range.End.Row, 1]);
            s4.HeaderAddress = ws.Cells["J20"];

            var s5 = chart.Series.Add(ws.Cells[21, 11, range.End.Row, 11], ws.Cells[21, 1, range.End.Row, 1]);
            s5.HeaderAddress = ws.Cells["K20"];

            chart.XAxis.Crosses = eCrosses.Min;
            chart.To.Row = 19;
            chart.To.Column = 11;
            chart.Legend.Add();
            chart.Legend.Position = eLegendPosition.Bottom;
            chart.StyleManager.SetChartStyle(chartStyle);
            chart.XAxis.TextBody.Rotation = -1000; //This is Excels "hack" to set the labels to auto rotation (horizontal/diagonal or vertical depending on space).
        }
    }
}
