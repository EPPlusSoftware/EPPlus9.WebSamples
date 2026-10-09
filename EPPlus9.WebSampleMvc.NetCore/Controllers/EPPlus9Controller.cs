using EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Fonts;
using System.IO;
using System.Threading.Tasks;

namespace EPPlus9.WebSampleMvc.NetCore.Controllers
{
    public class EPPlus9Controller : Controller
    {
        IWebHostEnvironment _env;
        public EPPlus9Controller(IWebHostEnvironment env)
        {
            _env = env;
        }
        private const string ContentTypeExcel = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private const string ContentTypePdf = "application/pdf";

        public IActionResult Index()
        {
            return View();
        }
        public async Task<IActionResult> HtmlExportwithCharts(HtmlExportWithChartModel model, string action)
        {
            if (action == "download")
            {
                return File(HtmlExportWithChartModel.CreateWorkbook(model.ChartStyle, model.TableStyle).GetAsByteArray(), ContentTypeExcel);
            }
            await model.LoadHtmlAndChartExport(model.ChartStyle, model.TableStyle);            
            return View(model);
        }
        public async Task<IActionResult> HtmlExportDifferentCharts(HtmlExportDifferentChartsModel model, string action)
        {
            if(action=="download")
            {
                return File(HtmlExportDifferentChartsModel.CreateWorkbook(model.ChartType, model.ChartStyle, model.TableStyle).GetAsByteArray(), ContentTypeExcel);
            }
            await model.LoadHtmlAndChartExport();
            return View(model);
        }
        public async Task<IActionResult> PdfExportTable(PdfExportTableModel model, string action)
        {
            if (action == "pdf")
            {
                using var pck = model.CreateWorkbook(_env.WebRootPath, model);
                ConfigureSampleFonts(pck.Workbook);
                using var ms = new MemoryStream();
                pck.Workbook.SaveAsPdf(ms);
                return File(ms.ToArray(), ContentTypePdf, "EPPlus Sample 3.pdf");
            }
            if (action=="excel")
            {
                using var pck = model.CreateWorkbook(_env.WebRootPath, model);
                return File(pck.GetAsByteArray(), ContentTypeExcel, "EPPlus Sample 3.xlsx");
            }
            await model.LoadHtml(_env.WebRootPath, model);
            return View(model);
        }

        private void ConfigureSampleFonts(ExcelWorkbook workbook, IFontLogger logger = null)
        {
            workbook.ConfigureFonts(cfg =>
            {
                cfg.FontDirectories.Add(Path.Combine(_env.ContentRootPath, "data", "Fonts"));
                cfg.SearchSystemDirectories = false;
                if (logger != null)
                {
                    cfg.Logger = logger;
                }
            });
        }

        [HttpGet]
        public IActionResult PdfExportTableFonts()
        {
            var model = new PdfExportTableModel();
            model.LoadFonts(Path.Combine(_env.ContentRootPath, "data", "Fonts"));
            return PartialView("_PdfExportTableFonts", model);
        }

        [HttpPost]
        public IActionResult PdfExportTableLog(PdfExportTableModel model, FontLogSeverity minSeverity = FontLogSeverity.Information)
        {
            var logger = new InMemoryFontLogger(minSeverity);
            using var pck = model.CreateWorkbook(_env.WebRootPath, model);
            ConfigureSampleFonts(pck.Workbook, logger);
            using var ms = new MemoryStream();
            pck.Workbook.SaveAsPdf(ms); // The PDF is discarded, we only want the log
            var logModel = new PdfExportTableLogModel(logger.GetSnapshot(), minSeverity, logger.DroppedCount, ms.Length);
            return PartialView("_PdfExportTableLog", logModel);
        }

        public async Task<IActionResult> PdfExportRange(PdfExportRangeModel model, string action)
        {
            if (action == "pdf")
            {
                using var pck = model.CreateWorkbook(_env.ContentRootPath);
                using var ms = new MemoryStream();
                pck.Workbook.SaveAsPdf(ms);
                return File(ms.ToArray(), ContentTypePdf, "EPPlus Sample 4.pdf");
            }
            if (action == "excel")
            {
                using var pck = model.CreateWorkbook(_env.ContentRootPath);
                return File(pck.GetAsByteArray(), ContentTypeExcel, "EPPlus Sample 4.xlsx");
            }
            await model.LoadHtml(_env.ContentRootPath);
            return View(model);
        }

        [HttpGet]
        public async Task<IActionResult> HtmlExportShapesToSvg()
        {
            var model = new HtmlExportShapesToSvgModel();
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> HtmlExportShapesToSvg(HtmlExportShapesToSvgModel model, string shape, string action)
        {
            var package = model.CreateWorkbookWithShape(shape, action);
            model.SelectedShapeName = shape;
            if (action == "excel")
            {
                return File(package.GetAsByteArray(), ContentTypeExcel, "EPPlus Sample 5.xlsx");
            }
            return View(model);
        }
    }
}
