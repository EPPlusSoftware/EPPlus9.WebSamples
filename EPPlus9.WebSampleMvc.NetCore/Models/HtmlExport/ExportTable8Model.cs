using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Globalization;
using System.IO;

namespace EPPlus9.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportTable8Model
    {
        public void SetupSampleData(TableStyles style = TableStyles.Dark3)
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Html export sample 8");
                var csvFileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\currencies2011weekly.csv"));
                var format = new ExcelTextFormat
                {
                    Delimiter = ';',
                    Culture = CultureInfo.InvariantCulture,
                    DataTypes = new eDataTypes[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number }
                };
                var tableRange = sheet.Cells["A15"].LoadFromText(csvFileInfo, format, style, true);

                sheet.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[tableRange.Start.Row, 1, tableRange.End.Row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                sheet.Cells[tableRange.Start.Row, 2, tableRange.End.Row, 5].Style.Numberformat.Format = "#,##0.0000";
                tableRange.AutoFitColumns();

                var table = sheet.Tables.GetFromRange(tableRange);
                table.ShowFirstColumn = true;
                var chart = sheet.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
                chart.Series.Add(tableRange.TakeColumns(1), tableRange.TakeColumns(2));
                chart.SetPosition(0, 0);
                chart.To.Row = 15;
                chart.To.Column = 10;
                chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle5);
                
                var exporter = table.CreateHtmlExporter();
                var settings = exporter.Settings;
                settings.Drawings.Include = OfficeOpenXml.Export.HtmlExport.eDrawingInclude.Include;

                settings.Culture = CultureInfo.InvariantCulture;
                settings.TableId = "currency-table";
                settings.AdditionalTableClassNames.Add("table");
                settings.AdditionalTableClassNames.Add("table-sm");
                settings.AdditionalTableClassNames.Add("table-borderless");

                // export css and html
                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }
    }
}
