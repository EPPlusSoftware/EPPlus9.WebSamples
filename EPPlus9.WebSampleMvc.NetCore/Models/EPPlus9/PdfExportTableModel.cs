using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Composition;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9
{
    public class PdfExportTableModel
    {
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
        public bool ShowFirstColumn { get; set; }

        public bool ShowLastColumn { get; set; }

        public bool ShowColumnStripes { get; set; }

        public bool ShowRowsStripes { get; set; }

        public TableStyles TableStyle { get; set; } = TableStyles.Dark3;

        public string TableStyleName => TableStyle.ToString();

        public string Html { get; set; }
        public string Css { get; set; }
        public ExcelPackage CreateWorkbook(string webRootPath, bool showFirstColumn, bool showLastColumn, bool showColumnStripes, bool showRowStripes)
        {
            InitDataTable();
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Pdf export sample 4");
            var tableRange = sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, TableStyle);

            //Configure the table
            var table = sheet.Tables.GetFromRange(tableRange);
            table.Sort(x => x.SortBy.ColumnNamed("Population", eSortOrder.Descending));
            table.ShowTotal = true;
            table.Columns[0].TotalsRowLabel = "Total";
            table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            table.Columns[2].TotalsRowFunction = RowFunctions.Sum;

            //Add column for population density
            table.Columns.Add(1);
            tableRange = table.Range;
            table.Columns[3].CalculatedColumnFormula = $"{table.Name}[[#This Row],[Population]]/{table.Name}[[#This Row],[Area (km²)]]";
            table.Columns[3].Name = "Density";
            table.Columns[3].TotalsRowFunction = RowFunctions.Average;
            table.ShowFirstColumn = showFirstColumn;
            table.ShowLastColumn = showLastColumn;
            table.ShowColumnStripes = showColumnStripes;
            table.ShowRowStripes = showRowStripes;
            sheet.Calculate();

            //// format the header
            table.Range.TakeColumnsBetween(1, 3).SkipRows(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            // format the rows
            var lastDataRow = tableRange.End.Row - 1;
            sheet.Cells[tableRange.Start.Row, 1, lastDataRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[tableRange.Start.Row, 2, lastDataRow, 2].Style.Numberformat.Format = "#,##0";
            sheet.Cells[tableRange.Start.Row, 3, lastDataRow, 3].Style.Numberformat.Format = "#,##0 \"km²\"";
            sheet.Cells[tableRange.Start.Row, 4, lastDataRow, 4].Style.Numberformat.Format = "#,##0.0";

            // format the total row
            var totalRow = tableRange.End.Row;
            sheet.Cells[totalRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[totalRow, 2].Style.Numberformat.Format = "#,##0";
            sheet.Cells[totalRow, 3].Style.Numberformat.Format = "#,##0 \"km²\"";
            sheet.Cells[totalRow, 4].Style.Numberformat.Format = "\"Avg: \"#,##0.0 ";
            sheet.Cells.AutoFitColumns();

            //Set the header and footer values
            SetHeader(sheet, webRootPath);
            SetFooter(sheet);

            sheet.PrinterSettings.TopMargin = 1.1;

            return package;
        }

        private void SetHeader(ExcelWorksheet sheet, string webRootPath)
        {
            var oddHeader = sheet.HeaderFooter.OddHeader;

            var imageFile = Path.Combine(webRootPath, "img", "EPPlus-logo-small.png");
            if (File.Exists(imageFile))
            {
                oddHeader.LeftAligned.AddImage(new FileInfo(imageFile));
            }

            var title = oddHeader.RightAligned.AddText("Population & Area Report");
            title.Bold = true;
            title.FontSize = 9;
            title.FontName = "Calibri";
        }

        private void SetFooter(ExcelWorksheet sheet)
        {
            var oddFooter = sheet.HeaderFooter.OddFooter;

            var footerLeft = oddFooter.LeftAligned.AddText("EPPlus Software AB");
            footerLeft.FontSize = 8;

            var pageLabel = oddFooter.Centered.AddText("Page ");
            pageLabel.FontSize = 9;
            var pageNumber = oddFooter.Centered.AddPageNumber();
            pageNumber.FontSize = 9;
            var ofLabel = oddFooter.Centered.AddText(" of ");
            ofLabel.FontSize = 9;
            var pageCount = oddFooter.Centered.AddNumberOfPages();
            pageCount.FontSize = 9;

            var footerRight = oddFooter.RightAligned.AddText("Confidential  •  Generated ");
            footerRight.Italic = true;
            footerRight.FontSize = 8;

            var dateItem = oddFooter.RightAligned.AddCurrentDate();
            dateItem.Italic = true;
            dateItem.FontSize = 8;
        }
        private DataTable _dataTable=null;

        private void InitDataTable()
        {
            if (_dataTable != null) return;
            _dataTable = new DataTable();
            _dataTable.Columns.Add("Country", typeof(string));
            _dataTable.Columns.Add("Population", typeof(int));
            var areaCol = _dataTable.Columns.Add("Area", typeof(int));
            areaCol.Caption = "Area (km²)";


            _dataTable.Rows.Add("Sweden", 10409248, 450295);
            _dataTable.Rows.Add("Norway", 5402171, 385178);
            _dataTable.Rows.Add("Netherlands", 17553530, 41198);
            _dataTable.Rows.Add("Finland", 5541806, 338145);
            _dataTable.Rows.Add("Belgium", 11521238, 30510);
            _dataTable.Rows.Add("Denmark", 5850189, 44493);
            _dataTable.Rows.Add("Lithuania", 2801264, 65300);
            _dataTable.Rows.Add("Greece", 10718565, 131940);
            _dataTable.Rows.Add("Russia", 145734038, 3972400);
            _dataTable.Rows.Add("Germany", 83124418, 357386);
            _dataTable.Rows.Add("France", 64990511, 551695);
            _dataTable.Rows.Add("Czech Republic", 10665677, 78866);
            _dataTable.Rows.Add("Slovakia", 5459781, 49036);
            _dataTable.Rows.Add("Spain", 47394223, 498468);
            _dataTable.Rows.Add("Portugal", 10256193, 91568);
            _dataTable.Rows.Add("United Kingdom", 67141684, 242495);
            _dataTable.Rows.Add("Poland", 37921592, 312685);
            _dataTable.Rows.Add("Albania", 2882740, 28748);
            _dataTable.Rows.Add("Estonia", 1322920, 45339);
            _dataTable.Rows.Add("Hungary", 9707499, 93030);
            _dataTable.Rows.Add("Romania", 19186000, 238397);
            _dataTable.Rows.Add("Italy", 60627291, 301338);
            _dataTable.Rows.Add("Bulgaria", 7051608, 110994);
            _dataTable.Rows.Add("Belarus", 9452617, 207600);
            _dataTable.Rows.Add("Austria", 8891388, 83858);
            _dataTable.Rows.Add("Switzerland", 8525611, 41290);
            _dataTable.Rows.Add("Ireland", 4818690, 70273);
            _dataTable.Rows.Add("Ukraine", 44246156, 603628);
            _dataTable.Rows.Add("Iceland", 336713, 102775);
            _dataTable.Rows.Add("Serbia", 6871547, 77453);
            _dataTable.Rows.Add("Croatia", 4156405, 56594);
            _dataTable.Rows.Add("Latvia", 1928459, 64589);
            _dataTable.Rows.Add("Bosnia and Herzegovina", 3323925, 51129);
            _dataTable.Rows.Add("Montenegro", 627809, 13812);
            _dataTable.Rows.Add("Cyrprus", 1189265, 9251);
            _dataTable.Rows.Add("Kosovo", 1798506, 10908);
            _dataTable.Rows.Add("Slovenia", 2055496, 20273);
            _dataTable.Rows.Add("Moldova", 4033963, 33846);
            _dataTable.Rows.Add("North Macedonia", 2083374, 25713);
            _dataTable.Rows.Add("United States", 331002651, 9833517);
            _dataTable.Rows.Add("China", 1412600000, 9596961);
            _dataTable.Rows.Add("India", 1417173173, 3287263);
            _dataTable.Rows.Add("Japan", 125681593, 377975);
            _dataTable.Rows.Add("Brazil", 214326223, 8515767);
            _dataTable.Rows.Add("Canada", 38246108, 9984670);
            _dataTable.Rows.Add("Australia", 25788215, 7692024);
            _dataTable.Rows.Add("Mexico", 126014024, 1964375);
            _dataTable.Rows.Add("South Africa", 59308690, 1221037);
            _dataTable.Rows.Add("Egypt", 104258327, 1001450);
            _dataTable.Rows.Add("Nigeria", 218541212, 923768);
            _dataTable.Rows.Add("Argentina", 45808747, 2780400);
            _dataTable.Rows.Add("Indonesia", 273523615, 1904569);
            _dataTable.Rows.Add("South Korea", 51780579, 100210);
            _dataTable.Rows.Add("Turkey", 84339067, 783562);
            _dataTable.Rows.Add("Saudi Arabia", 34813871, 2149690);
            _dataTable.Rows.Add("Israel", 9053300, 20770);
            _dataTable.Rows.Add("New Zealand", 5084300, 268838);
            _dataTable.Rows.Add("Thailand", 69950850, 513120);
            _dataTable.Rows.Add("Vietnam", 98168833, 331212);
        }

        internal async Task LoadHtml(string webRootPath, bool showFirstColumn, bool showLastColumn, bool showColumnStripes, bool showRowStripes)
        {
            var package = CreateWorkbook(webRootPath, showFirstColumn, showLastColumn, showColumnStripes, showRowStripes);
            
            var sheet = package.Workbook.Worksheets[0];
            var exporter = sheet.Cells.CreateHtmlExporter();

            var settings = exporter.Settings;
            settings.Culture = CultureInfo.InvariantCulture;
            settings.TableId = "countries-table";
            settings.AdditionalTableClassNames.Add("table");
            settings.AdditionalTableClassNames.Add("table-sm");
            settings.AdditionalTableClassNames.Add("table-borderless");

            Html = await exporter.GetHtmlStringAsync(); // Get the HTML string for the worksheet without the chart as a table.
            Css = await exporter.GetCssStringAsync();   //Get the CSS string for the worksheet without the chart as a table. You could include the drawing in the css as a class, but in this case we want to export the chart as a separate SVG.
        }
    }
}
