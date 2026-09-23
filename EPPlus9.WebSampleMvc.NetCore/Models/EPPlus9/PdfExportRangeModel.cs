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
    public class PdfExportRangeModel
    {
        public string Html { get; set; }
        public string Css { get; set; }
        public ExcelPackage CreateWorkbook(string webRootPath)
        {
            var file = Path.Combine(webRootPath, "data\\Allsvenskan2001.xlsx");
            return new ExcelPackage(file);
        }
        

        internal async Task LoadHtml(string webRootPath)
        {
            var package = CreateWorkbook(webRootPath);

            var sheet = package.Workbook.Worksheets[0];

            var exporter = sheet.Cells["B5:N19"].CreateHtmlExporter();
            var settings = exporter.Settings;

            settings.Culture = CultureInfo.InvariantCulture;
            settings.TableId = "soccer-table";
            settings.Accessibility.TableSettings.AriaLabel = "This html-table is a range that is exported from EPPlus";

            // use column width from the workbook
            settings.SetColumnWidth = true;
            settings.SetRowHeight = true;

            // include pictures in the exported cells
            settings.Drawings.Include = eDrawingInclude.Include;
            settings.Drawings.AddNameAsId = false;

            // when Minify is false the output will be formated with indentation and linebreaks.
            settings.Minify = false;

            // export css and html
            Css = exporter.GetCssString();
            Html = exporter.GetHtmlString();
        }
    }
}
