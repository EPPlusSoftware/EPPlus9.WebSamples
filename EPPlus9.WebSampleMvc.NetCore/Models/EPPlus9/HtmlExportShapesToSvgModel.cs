using NuGet.ContentModel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;

namespace EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9
{
    public class HtmlExportShapesToSvgModel
    {
        public IEnumerable<eShapeStyle> AllShapes
        {
            get
            {
                var list = new List<eShapeStyle>();
                foreach(var shapeItem in Enum.GetValues(typeof(eShapeStyle)))
                {
                    var shapeStyle = (eShapeStyle)shapeItem;
                    if (shapeStyle == eShapeStyle.CustomShape) continue;
                    list.Add(shapeStyle);
                }
                return list;
            }
        }

        public string SelectedShapeName { get; set; }

        public string Svg 
        {
            get; set;
        }


        public ExcelPackage CreateWorkbookWithShape(string shapeStyleName, string action)
        {
            Enum.TryParse(shapeStyleName, out eShapeStyle shapeStyle);
            var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Shapes to SVG");
            var shape = ws.Drawings.AddShape(shapeStyleName, shapeStyle);
            shape.Text = shapeStyleName;
            shape.SetPosition(100, 100);
            shape.SetSize(600, 600);
            if(action != "excel0")
            {
                Svg = shape.ToSvg();
            }
            return package;
        }
    }
}