using EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
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
            if(action=="pdf")
            {
                using var pck = model.CreateWorkbook(_env.WebRootPath, model.ShowFirstColumn, model.ShowLastColumn, model.ShowColumnStripes, model.ShowRowsStripes);
                using  var ms = new MemoryStream();
                pck.Workbook.SaveAsPdf(ms);
                return File(ms.ToArray(), ContentTypePdf, "EPPlus Sample 3.pdf");
            }
            if(action=="excel")
            {
                using var pck = model.CreateWorkbook(_env.WebRootPath, model.ShowFirstColumn, model.ShowLastColumn, model.ShowColumnStripes, model.ShowRowsStripes);
                return File(pck.GetAsByteArray(), ContentTypeExcel, "EPPlus Sample 3.xlsx");
            }
            model.ShowRowsStripes = true;
            await model.LoadHtml(_env.WebRootPath, model.ShowFirstColumn, model.ShowLastColumn, model.ShowColumnStripes, model.ShowRowsStripes);
            return View(model);
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
    }
}
