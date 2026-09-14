using EPPlus.Export.Pdf.Enums;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Composition;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9
{
    public class HtmlExportDifferentChartsModel
    {
        public class RegionalSales
        {
            public string Region { get; set; }
            public int SoldUnits { get; set; }
            public double TotalSales { get; set; }
            public double Margin { get; set; }
        }
        public enum SelectedChartType
        {
            ComboChart,
            LineChart,
            ColumnChart,
            BarChart,
            PieChart
        }

        private static List<RegionalSales> _salesData = new List<RegionalSales>()
        {
                new RegionalSales(){ Region = "North", SoldUnits=500, TotalSales=4800, Margin=0.200 },
                new RegionalSales(){ Region = "Central", SoldUnits=900, TotalSales=7330, Margin=0.333 },
                new RegionalSales(){ Region = "South", SoldUnits=400, TotalSales=3700, Margin=0.150 },
                new RegionalSales(){ Region = "East", SoldUnits=350, TotalSales=4400, Margin=0.102 },
                new RegionalSales(){ Region = "West", SoldUnits=700, TotalSales=6900, Margin=0.218 },
                new RegionalSales(){ Region = "Stockholm", SoldUnits=1200, TotalSales=8250, Margin=0.350 }
        };

        public SelectedChartType ChartType { get; set; } = SelectedChartType.ColumnChart;
        public ePresetChartStyleMultiSeries ChartStyle { get; set; } = ePresetChartStyleMultiSeries.ColumnChartStyle1;
        public TableStyles TableStyle { get; set; } = TableStyles.Dark3;
        public IEnumerable<SelectListItem> SupportedChartTypes
        {
            get
            {
                return Enum.GetValues(typeof(SelectedChartType))
                    .Cast<SelectedChartType>()
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> AllChartStyles
        { 
            get 
            {
                switch (ChartType)
                {
                    case SelectedChartType.ComboChart:
                        return [new SelectListItem("Style 1", ePresetChartStyleMultiSeries.ComboChartStyle1.ToString()),
                                new SelectListItem("Style 2", ePresetChartStyleMultiSeries.ComboChartStyle2.ToString()),
                                new SelectListItem("Style 3", ePresetChartStyleMultiSeries.ComboChartStyle3.ToString()),
                                new SelectListItem("Style 4", ePresetChartStyleMultiSeries.ComboChartStyle4.ToString()),
                                new SelectListItem("Style 5", ePresetChartStyleMultiSeries.ComboChartStyle5.ToString()),
                                new SelectListItem("Style 6", ePresetChartStyleMultiSeries.ComboChartStyle6.ToString()),
                                new SelectListItem("Style 7", ePresetChartStyleMultiSeries.ComboChartStyle7.ToString()),
                                new SelectListItem("Style 8", ePresetChartStyleMultiSeries.ComboChartStyle8.ToString()),
                               ];
                    case SelectedChartType.LineChart:
                        return [new SelectListItem("Style 1", ePresetChartStyleMultiSeries.LineChartStyle1.ToString()),
                                new SelectListItem("Style 2", ePresetChartStyleMultiSeries.LineChartStyle2.ToString()),
                                new SelectListItem("Style 3", ePresetChartStyleMultiSeries.LineChartStyle3.ToString()),
                                new SelectListItem("Style 4", ePresetChartStyleMultiSeries.LineChartStyle4.ToString()),
                                new SelectListItem("Style 5", ePresetChartStyleMultiSeries.LineChartStyle5.ToString()),
                                new SelectListItem("Style 6", ePresetChartStyleMultiSeries.LineChartStyle6.ToString()),
                                new SelectListItem("Style 7", ePresetChartStyleMultiSeries.LineChartStyle7.ToString()),
                                new SelectListItem("Style 8", ePresetChartStyleMultiSeries.LineChartStyle8.ToString()),
                                new SelectListItem("Style 9", ePresetChartStyleMultiSeries.LineChartStyle9.ToString()),
                                new SelectListItem("Style 10", ePresetChartStyleMultiSeries.LineChartStyle10.ToString()),
                                new SelectListItem("Style 11", ePresetChartStyleMultiSeries.LineChartStyle11.ToString()),
                                new SelectListItem("Style 12", ePresetChartStyleMultiSeries.LineChartStyle12.ToString()),
                                new SelectListItem("Style 13", ePresetChartStyleMultiSeries.LineChartStyle13.ToString()),
                               ];
                    case SelectedChartType.BarChart:
                        return [new SelectListItem("Style 1", ePresetChartStyleMultiSeries.BarChartStyle1.ToString()),
                                new SelectListItem("Style 2", ePresetChartStyleMultiSeries.BarChartStyle2.ToString()),
                                new SelectListItem("Style 3", ePresetChartStyleMultiSeries.BarChartStyle3.ToString()),
                                new SelectListItem("Style 4", ePresetChartStyleMultiSeries.BarChartStyle4.ToString()),
                                new SelectListItem("Style 5", ePresetChartStyleMultiSeries.BarChartStyle5.ToString()),
                                new SelectListItem("Style 6", ePresetChartStyleMultiSeries.BarChartStyle6.ToString()),
                                new SelectListItem("Style 7", ePresetChartStyleMultiSeries.BarChartStyle7.ToString()),
                                new SelectListItem("Style 8", ePresetChartStyleMultiSeries.BarChartStyle8.ToString()),
                                new SelectListItem("Style 9", ePresetChartStyleMultiSeries.BarChartStyle9.ToString()),
                                new SelectListItem("Style 10", ePresetChartStyleMultiSeries.BarChartStyle10.ToString()),
                                new SelectListItem("Style 11", ePresetChartStyleMultiSeries.BarChartStyle11.ToString()),
                                new SelectListItem("Style 12", ePresetChartStyleMultiSeries.BarChartStyle12.ToString()),
                               ];
                    case SelectedChartType.ColumnChart:
                        return [new SelectListItem("Style 1", ePresetChartStyleMultiSeries.ColumnChartStyle1.ToString()),
                                new SelectListItem("Style 2", ePresetChartStyleMultiSeries.ColumnChartStyle2.ToString()),
                                new SelectListItem("Style 3", ePresetChartStyleMultiSeries.ColumnChartStyle3.ToString()),
                                new SelectListItem("Style 4", ePresetChartStyleMultiSeries.ColumnChartStyle4.ToString()),
                                new SelectListItem("Style 5", ePresetChartStyleMultiSeries.ColumnChartStyle5.ToString()),
                                new SelectListItem("Style 6", ePresetChartStyleMultiSeries.ColumnChartStyle6.ToString()),
                                new SelectListItem("Style 7", ePresetChartStyleMultiSeries.ColumnChartStyle7.ToString()),
                                new SelectListItem("Style 8", ePresetChartStyleMultiSeries.ColumnChartStyle8.ToString()),
                                new SelectListItem("Style 9", ePresetChartStyleMultiSeries.ColumnChartStyle9.ToString()),
                                new SelectListItem("Style 10", ePresetChartStyleMultiSeries.ColumnChartStyle10.ToString()),
                                new SelectListItem("Style 11", ePresetChartStyleMultiSeries.ColumnChartStyle11.ToString()),
                                new SelectListItem("Style 12", ePresetChartStyleMultiSeries.ColumnChartStyle12.ToString()),
                                new SelectListItem("Style 13", ePresetChartStyleMultiSeries.ColumnChartStyle13.ToString()),
                                new SelectListItem("Style 14", ePresetChartStyleMultiSeries.ColumnChartStyle14.ToString()),
                               ];
                    case SelectedChartType.PieChart:
                    default:
                        return [new SelectListItem("Style 1", ePresetChartStyleMultiSeries.PieChartStyle1.ToString()),
                                new SelectListItem("Style 2", ePresetChartStyleMultiSeries.PieChartStyle2.ToString()),
                                new SelectListItem("Style 3", ePresetChartStyleMultiSeries.PieChartStyle3.ToString()),
                                new SelectListItem("Style 4", ePresetChartStyleMultiSeries.PieChartStyle4.ToString()),
                                new SelectListItem("Style 5", ePresetChartStyleMultiSeries.PieChartStyle5.ToString()),
                                new SelectListItem("Style 6", ePresetChartStyleMultiSeries.PieChartStyle6.ToString()),
                                new SelectListItem("Style 7", ePresetChartStyleMultiSeries.PieChartStyle7.ToString()),
                                new SelectListItem("Style 8", ePresetChartStyleMultiSeries.PieChartStyle8.ToString()),
                                new SelectListItem("Style 9", ePresetChartStyleMultiSeries.PieChartStyle9.ToString()),
                                new SelectListItem("Style 10", ePresetChartStyleMultiSeries.PieChartStyle10.ToString()),
                                new SelectListItem("Style 11", ePresetChartStyleMultiSeries.PieChartStyle11.ToString()),
                                new SelectListItem("Style 12", ePresetChartStyleMultiSeries.PieChartStyle12.ToString()),
                               ];
                }
            }
        }


        public IEnumerable<SelectListItem> AllBuiltInTableStyles
        {
            get
            {
                return Enum.GetValues(typeof(TableStyles))
                    .Cast<TableStyles>()
                    .Where(x => x != TableStyles.Custom)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public string Html { get; private set; }
        public string Css { get; private set; }
        public string SvgChart { get; private set; }

        public async Task LoadHtmlAndChartExport()
        {
            var package = CreateWorkbook(ChartType, ChartStyle, TableStyle);
            var sheet = package.Workbook.Worksheets[0];
            var exporter = sheet.Cells.CreateHtmlExporter();
            var settings = exporter.Settings;
            settings.Drawings.Include = eDrawingInclude.IncludeInCssOnly; // Exclude the chart from the HTML export, we will export it as a separate SVG. Optinally you could use the eDrawingInclude.IncludeAsClass option to include the chart in the HTML export as a class.
            settings.Culture = CultureInfo.InvariantCulture;

            settings.TableId = "currency-table";
            settings.AdditionalTableClassNames.Add("table");
            settings.AdditionalTableClassNames.Add("table-sm");
            settings.AdditionalTableClassNames.Add("table-borderless");
            
            Html = await exporter.GetHtmlStringAsync(); // Get the HTML string for the worksheet without the chart as a table.
            Css = await exporter.GetCssStringAsync();   //Get the CSS string for the worksheet without the chart as a table. You could include the drawing in the css as a class, but in this case we want to export the chart as a separate SVG.
            SvgChart = sheet.Drawings[0].ToSvg();       //Get the chart as a separate SVG string.
        }
        
        public static ExcelPackage CreateWorkbook(SelectedChartType chartType, ePresetChartStyleMultiSeries chartStyle, TableStyles tableStyle)
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Html export with svg chart");

            var range = sheet.Cells["A1"].LoadFromCollection(_salesData, true, tableStyle);
            sheet.Cells["B2:C7"].Style.Numberformat.Format = "#,##0";
            sheet.Cells["D2:D7"].Style.Numberformat.Format = "#,##0.00%";
            ExcelChart chart;
            switch (chartType)
            {
                case SelectedChartType.ComboChart:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.ColumnClustered);
                    var s1 = chart.Series.Add(sheet.Cells["B2:B7"], sheet.Cells["A2:A7"]);
                    s1.HeaderAddress = sheet.Cells["B1"];
                    var s2 = chart.Series.Add(sheet.Cells["C2:C7"], sheet.Cells["A2:A7"]);
                    s2.HeaderAddress = sheet.Cells["C1"];
                    var lineChartType = chart.PlotArea.ChartTypes.Add(eChartType.Line);

                    lineChartType.UseSecondaryAxis = true;
                    var s3 = lineChartType.Series.Add(sheet.Cells["D2:D7"], sheet.Cells["A2:A7"]);
                    s3.HeaderAddress = sheet.Cells["D1"];
                    break;
                case SelectedChartType.LineChart:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.LineMarkers);
                    var s=chart.Series.Add(sheet.Cells["D2:D7"], sheet.Cells["A2:A7"]);
                    s.HeaderAddress = sheet.Cells["D1"];
                    break;
                case SelectedChartType.ColumnChart:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.ColumnClustered);
                    s1 = chart.Series.Add(sheet.Cells["B2:B7"], sheet.Cells["A2:A7"]);
                    s1.HeaderAddress = sheet.Cells["B1"];
                    s2 = chart.Series.Add(sheet.Cells["C2:C7"], sheet.Cells["A2:A7"]);
                    s2.HeaderAddress = sheet.Cells["C1"];
                    var barChart = chart as ExcelBarChart;
                    barChart.Overlap = -5;

                    break;
                case SelectedChartType.BarChart:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.BarClustered);
                    s1 = chart.Series.Add(sheet.Cells["B2:B7"], sheet.Cells["A2:A7"]);
                    s1.HeaderAddress = sheet.Cells["B1"];
                    s2 = chart.Series.Add(sheet.Cells["C2:C7"], sheet.Cells["A2:A7"]);
                    s2.HeaderAddress = sheet.Cells["C1"];
                    barChart = chart as ExcelBarChart;
                    barChart.Overlap = -10;
                    break;
                case SelectedChartType.PieChart:
                default:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.PieExploded);
                    chart.Series.Add(sheet.Cells["D2:D7"], sheet.Cells["A2:A7"]);
                    var pieChart = chart as ExcelPieChart;
                    chart.StyleManager.SetChartStyle(chartStyle);
                    pieChart.DataLabel.ShowPercent = true;
                    pieChart.DataLabel.Border.Fill.Style=eFillStyle.SolidFill;
                    pieChart.DataLabel.Border.Fill.Color = Color.Black;
                    pieChart.DataLabel.Fill.Style = eFillStyle.SolidFill;
                    pieChart.DataLabel.Fill.Color = Color.LightCoral;
                    chart.SetPosition(2, 0, 5, 0);
                    chart.SetSize(1100, 300);
                    return package;
            }

            chart.StyleManager.SetChartStyle(chartStyle);
            chart.SetPosition(2,0,5, 0);
            chart.SetSize(1100,400);

            return package;
        }

    }
}
