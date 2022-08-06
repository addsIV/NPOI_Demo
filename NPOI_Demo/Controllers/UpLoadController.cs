using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;

namespace NPOI_Demo.Controllers
{
	public class UpLoadController : Controller
	{
		// GET: UpLoad
		[HttpPost]
		public ActionResult Uploads(HttpPostedFileBase file)
		{
			//判斷是否有上傳檔案
			if (Request.Files["file"].ContentLength > 0)
			{
				string extension = System.IO.Path.GetExtension(file.FileName);
				string fileSavedPath = WebConfigurationManager.AppSettings["UploadPath"];

				//判斷是否為excel檔
				if (extension == ".xls" || extension == ".xlsx")
				{
					//更改檔名為當天日期時間
					string newFileName = string.Concat(
					DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"),
					Path.GetExtension(file.FileName).ToLower());
					string fullFilePath = Path.Combine(Server.MapPath(fileSavedPath), newFileName);

					// 存放檔案到伺服器上
					Request.Files["file"].SaveAs(fullFilePath);

					//將資料傳送到前端顯示
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

			return View("ExcelImport");
		}
	}
}