using EPPlus.Export.Pdf.Enums;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Composition;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9
{
    public class HtmlExportWithChartModel
    {
        public eChartType ChartType { get; set; } = eChartType.Line;
        public ePresetChartStyleMultiSeries ChartStyle { get; set; } = ePresetChartStyleMultiSeries.LineChartStyle5;
        public TableStyles TableStyle { get; set; } = TableStyles.Dark3;
        public IEnumerable<SelectListItem> AllChartStyles
        { 
            get 
            {
                return [
                    new SelectListItem("Style 1", ePresetChartStyleMultiSeries.LineChartStyle1.ToString()),
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
                    //new SelectListItem("Style 14", ePresetChartStyleM.LineChartStyle14.ToString()),
                    //new SelectListItem("Style 15", ePresetChartStyle.LineChartStyle15.ToString())
                        ];
            } 
        }
        
        public IEnumerable<SelectListItem> AllBuiltInTableStyles
        {
            get
            {
                return System.Enum.GetValues(typeof(TableStyles))
                    .Cast<TableStyles>()
                    .Where(x => x != TableStyles.Custom)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public string Html { get; private set; }
        public string Css { get; private set; }     

        public async Task LoadHtmlAndChartExport(ePresetChartStyleMultiSeries chartStyle, TableStyles tableStyle)
        {
            var package = CreateWorkbook(chartStyle, tableStyle);
            var sheet = package.Workbook.Worksheets[0];
            var exporter = sheet.Cells.CreateHtmlExporter();
            var settings = exporter.Settings;
            settings.Drawings.Include = eDrawingInclude.Include;
            settings.Culture = CultureInfo.InvariantCulture;
            
            settings.Drawings.Position = eDrawingPosition.Absolute;
            settings.SetRowHeight = true;
            settings.SetColumnWidth = true;
            
            settings.TableId = "currency-table";
            
            Html = await exporter.GetHtmlStringAsync();
            Css = await exporter.GetCssStringAsync();   
        }

        public static ExcelPackage CreateWorkbook(ePresetChartStyleMultiSeries chartStyle, TableStyles tableStyle)
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Html export with svg chart");
            var csvFileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\currencies2011weekly.csv"));
            if (csvFileInfo.Exists == false) return null;
            var format = new ExcelTextFormat
            {
                Delimiter = ';',
                Culture = CultureInfo.InvariantCulture,
                DataTypes = new eDataTypes[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number }
            };
            var tableRange = sheet.Cells["A15"].LoadFromText(csvFileInfo, format, tableStyle, true);

            sheet.Cells["B1:E1"].Style.HorizontalAlignment = (OfficeOpenXml.Style.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right;
            sheet.Cells[tableRange.Start.Row, 1, tableRange.End.Row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
            sheet.Cells[tableRange.Start.Row, 2, tableRange.End.Row, 5].Style.Numberformat.Format = "#,##0.0000";
            tableRange.AutoFitColumns();

            var table = sheet.Tables.GetFromRange(tableRange);
            table.ShowFirstColumn = true;
            var chart = sheet.Drawings.AddLineChart("LineChart1", eLineChartType.Line);

            var serie1 = chart.Series.Add(tableRange.TakeColumnsBetween(1, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
            serie1.HeaderAddress = sheet.Cells["B15"];

            var serie2 = chart.Series.Add(tableRange.TakeColumnsBetween(2, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
            serie2.HeaderAddress = sheet.Cells["C15"]; 

            var serie3 = chart.Series.Add(tableRange.TakeColumnsBetween(3, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
            serie3.HeaderAddress = sheet.Cells["D15"];

            var serie4 = chart.Series.Add(tableRange.TakeColumnsBetween(4, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
            serie4.HeaderAddress = sheet.Cells["E15"];


            chart.SetPosition(0, 0);
            chart.To.Row = 14;
            chart.To.Column = 10;
            chart.StyleManager.SetChartStyle(chartStyle);

            var textBox = sheet.Drawings.AddShape("InfoBox", eShapeStyle.Rect);
            textBox.RichText.Add("This is a line chart with data from the table below. The chart is exported as SVG when exporting to HTML.");
            textBox.SetPosition(2, 0, 10, 0);
            textBox.SetSize(300, 200);
            
            return package;
        }

    }
}
