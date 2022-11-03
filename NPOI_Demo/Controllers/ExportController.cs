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
        public ActionResult Exports(int firstCellValue, int secondCellValue, string thirdCellFormula)
        {
            using (var exportData = new MemoryStream())
            {
                var workbook = new HSSFWorkbook();
                var sheet = workbook.CreateSheet("Sheet1");

                sheet.CreateRow(0);
                sheet.GetRow(0).CreateCell(0).SetCellType(CellType.Numeric);
                sheet.GetRow(0).CreateCell(0).SetCellValue(firstCellValue);

                sheet.GetRow(0).CreateCell(1).SetCellType(CellType.Numeric);
                sheet.GetRow(0).CreateCell(1).SetCellValue(secondCellValue);

                sheet.GetRow(0).CreateCell(2).SetCellFormula(thirdCellFormula);

                workbook.Write(exportData);
                var saveAsFileName = $"NPOI demo-{DateTime.Now:d}.xls".Replace("/", "-");

                var bytes = exportData.ToArray();
                return File(bytes, "application/vnd.ms-excel", saveAsFileName);
            }
        }
    }
}