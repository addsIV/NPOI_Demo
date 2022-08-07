using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI_Demo.Controllers
{
	public class UpLoadController : Controller
	{
		[HttpPost]
		public ActionResult Uploads(HttpPostedFileBase file)
		{
			UploadService(file);

			ShowExcelValue();
			return View("ExcelImport");
		}

		private void ShowExcelValue()
		{
			//if (!ViewBag.path) return;

			IWorkbook workBook = new XSSFWorkbook(ViewBag.path);
			ISheet sheet = workBook.GetSheetAt(0);

			var formula = new XSSFFormulaEvaluator(workBook);
			var first_Cell = sheet.GetRow(0).GetCell(0);
			var second_Cell = sheet.GetRow(0).GetCell(1);


			ViewBag.First_Cell_Value = first_Cell.ToString();
			ViewBag.Second_Cell_Formula = second_Cell.ToString();
			formula.EvaluateFormulaCell(second_Cell);
			ViewBag.Second_Cell_Nermeric = second_Cell.NumericCellValue.ToString();

		}

		public void UploadService(HttpPostedFileBase file)
		{
			if (Request.Files["file"].ContentLength > 0)
			{
				string extension = System.IO.Path.GetExtension(file.FileName);
				string fileSavedPath = WebConfigurationManager.AppSettings["UploadPath"];

				if (extension == ".xls" || extension == ".xlsx")
				{
					string newFileName = string.Concat(
					DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"),
					Path.GetExtension(file.FileName).ToLower());
					string fullFilePath = Path.Combine(Server.MapPath(fileSavedPath), newFileName);

					Request.Files["file"].SaveAs(fullFilePath);

					ViewBag.path = fullFilePath;
					Response.Write("<script language=javascript>alert(' 檔案上傳成功 ');</" +
					"script>");
				}
				else
				{
					Response.Write(@"<script language=javascript>alert('請上傳 .xls  或 .xlsx 格式的檔 


					案');</" + "script>");

				}
			}
			else
			{
				Response.Write("<script language=javascript>alert('請選擇檔案');</" + "script>");
			}
		}
	}
}