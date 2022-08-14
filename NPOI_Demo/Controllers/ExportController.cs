using System;
using System.IO;
using System.Web.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOI_Demo.Controllers
{
    public class ExportController : Controller
    {
        [HttpPost]
        public ActionResult Exports(int First_Cell_Value, int Second_Cell_Value, string Third_Cell_Formula)
        {
            using (var exportData = new MemoryStream())
            {
                var workbook = new HSSFWorkbook();
                var sheet = workbook.CreateSheet("Sheet1");

                sheet.CreateRow(0);
                sheet.GetRow(0).CreateCell(0).SetCellType(CellType.Numeric);
                sheet.GetRow(0).CreateCell(0).SetCellValue(First_Cell_Value);

                sheet.GetRow(0).CreateCell(1).SetCellType(CellType.Numeric);
                sheet.GetRow(0).CreateCell(1).SetCellValue(Second_Cell_Value);

                sheet.GetRow(0).CreateCell(2).SetCellFormula(Third_Cell_Formula);

                workbook.Write(exportData);
                string saveAsFileName = string.Format("NPOI demo-{0:d}.xls", DateTime.Now).Replace("/", "-");

                byte[] bytes = exportData.ToArray();
                return File(bytes, "application/vnd.ms-excel", saveAsFileName);
            }
        }
    }
}