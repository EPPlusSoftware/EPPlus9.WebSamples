using EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Threading.Tasks;

namespace EPPlus9.WebSampleMvc.NetCore.Controllers
{
    public class EPPlus9Controller : Controller
    {
        private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public IActionResult Index()
        {
            return View();
        }
        public async Task<IActionResult> HtmlExportwithCharts(HtmlExportWithChartModel model)
        {
            await model.LoadHtmlAndChartExport(model.ChartStyle, model.TableStyle);            
            return View(model);
        }
        public async Task<IActionResult> HtmlExportDifferentCharts(HtmlExportDifferentChartsModel model, string action)
        {
            if(action=="download")
            {
                return File(HtmlExportDifferentChartsModel.CreateWorkbook(model.ChartType, model.ChartStyle, model.TableStyle).GetAsByteArray(), ContentType);
            }
            else
            {
                await model.LoadHtmlAndChartExport(model.ChartType, model.ChartStyle, model.TableStyle);
                return View(model);
            }
        }
        public async Task<IActionResult> DownloadWorkbook(HtmlExportWithChartModel model)
        {
            return File(HtmlExportWithChartModel.CreateWorkbook(model.ChartStyle, model.TableStyle).GetAsByteArray(), ContentType);
        }
        public async Task<IActionResult> DownloadDifferentChartWorkbook(HtmlExportDifferentChartsModel model)
        {
            return File(HtmlExportDifferentChartsModel.CreateWorkbook(model.ChartType, model.ChartStyle, model.TableStyle).GetAsByteArray(), ContentType);
        }
        public IActionResult DrawingsAsSvg()
        {
            return View();
        }
        public IActionResult PdfExport()
        {
            return View();
        }
    }
}
