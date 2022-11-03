using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.Json;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI_Demo.Models;

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
			var workBook = new XSSFWorkbook(ViewBag.path);
			var sheet = workBook.GetSheetAt(0);

			var formula = new XSSFFormulaEvaluator(workBook);
			var lastRowNum = sheet.LastRowNum;

			var missionAutos = new List<MissionAuto>();

			for (var i = 1; i < lastRowNum; i++)
			{
				var cell = sheet.GetRow(i).GetCell(0);
				missionAutos.Add(new MissionAuto()
				{
					Order = (int) cell.NumericCellValue,
					Group = (int) sheet.GetRow(i).GetCell(1).NumericCellValue,
					CriteriaType = sheet.GetRow(i).GetCell(2).StringCellValue,
					CriteriaAmount = (int) sheet.GetRow(i).GetCell(3).NumericCellValue,
					RewardType = sheet.GetRow(i).GetCell(4).StringCellValue,
					RewardProvider = sheet.GetRow(i).GetCell(5) == null?"":sheet.GetRow(i).GetCell(5).StringCellValue,
					RewardAmount = (int) sheet.GetRow(i).GetCell(6).NumericCellValue,
					PeriodInHours = (int) sheet.GetRow(i).GetCell(7).NumericCellValue,
				});
			}

			ViewBag.First_Cell_Value = JsonSerializer.Serialize(missionAutos);
			// ViewBag.Second_Cell_Formula = secondCell.ToString();
			// formula.EvaluateFormulaCell(secondCell);
			// ViewBag.Second_Cell_Nermeric = secondCell.NumericCellValue.ToString(CultureInfo.CurrentCulture);

		}

		public void UploadService(HttpPostedFileBase file)
		{
			if (Request.Files["file"].ContentLength > 0)
			{
				var extension = System.IO.Path.GetExtension(file.FileName);
				var fileSavedPath = WebConfigurationManager.AppSettings["UploadPath"];

				if (extension == ".xls" || extension == ".xlsx")
				{
					var newFileName = string.Concat(
					DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"),
					Path.GetExtension(file.FileName).ToLower());
					var fullFilePath = Path.Combine(Server.MapPath(fileSavedPath), newFileName);

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