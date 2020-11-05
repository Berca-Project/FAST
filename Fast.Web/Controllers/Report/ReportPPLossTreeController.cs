using Fast.Application.Interfaces;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Utils;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace Fast.Web.Controllers.Report
{
	[CustomAuthorize("ReportPPLossTree")]
	public class ReportPPLossTreeController : BaseController<PPLPHModel>
	{

		private readonly ILoggerAppService _logger;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IWeeksAppService _weeksAppService;
		private readonly IMachineAppService _machineAppService;
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		public ReportPPLossTreeController(

		 ILoggerAppService logger,
		 IReferenceAppService referenceAppService,
		 ILocationAppService locationAppService,
		 IEmployeeAppService employeeAppService,
		 IWeeksAppService weeksAppService,
		 IMachineAppService machineAppService,
		 IReferenceDetailAppService referenceDetailAppService)
		{

			_logger = logger;
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_weeksAppService = weeksAppService;
			_machineAppService = machineAppService;
			_referenceDetailAppService = referenceDetailAppService;
		}

		public ActionResult Index()
		{
			return View();
		}
		[HttpPost]
		public ActionResult ExportTable(Dictionary<string, string> data)
		{
			try
			{
				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Loss Tree");
				int rowPos = 1;


				Sheet.Cells[rowPos, 2].Value = data["area"];
				Sheet.Cells[rowPos, 2].Style.Font.Bold = true;

				rowPos = rowPos + 2;
				Sheet.Cells[rowPos, 2].Value = "From";
				Sheet.Cells[rowPos, 3].Value = data["startDate"] + " 06:00";
				Sheet.Cells[rowPos, 4].Value = "To";
				Sheet.Cells[rowPos, 5].Value = data["endDate"] + " 06:00";

				rowPos = rowPos + 2;
				Sheet.Cells[rowPos, 2].Value = "Calendar Time";
				Sheet.Cells[rowPos, 3].Value = data["CalendarTimeMin"];
				Sheet.Cells[rowPos, 4].Value = "Min";
				Sheet.Cells[rowPos, 5].Value = data["CalendarTimeHour"];
				Sheet.Cells[rowPos++, 6].Value = "Hour";

				Sheet.Cells[rowPos, 2].Value = "Exclude Time";
				Sheet.Cells[rowPos, 3].Value = data["ExcludeMinute"];
				Sheet.Cells[rowPos, 4].Value = "Min";
				Sheet.Cells[rowPos, 5].Value = data["ExcludeHour"];
				Sheet.Cells[rowPos++, 6].Value = "Hour";

				Sheet.Cells[rowPos, 2].Value = "Scheduled Time";
				Sheet.Cells[rowPos, 3].Value = data["ScheduledMinute"];
				Sheet.Cells[rowPos, 4].Value = "Min";
				Sheet.Cells[rowPos, 5].Value = data["ScheduledHour"];
				Sheet.Cells[rowPos++, 6].Value = "Hour";

				Sheet.Cells[rowPos, 2].Value = "Downtime";
				Sheet.Cells[rowPos, 3].Value = data["DowntimeMinute"];
				Sheet.Cells[rowPos, 4].Value = "Min";
				Sheet.Cells[rowPos, 5].Value = data["DowntimeHour"];
				Sheet.Cells[rowPos++, 6].Value = "Hour";

				Sheet.Cells[rowPos, 2].Value = "Running Time";
				Sheet.Cells[rowPos, 3].Value = data["RunningTimeMinute"];
				Sheet.Cells[rowPos, 4].Value = "Min";
				Sheet.Cells[rowPos, 5].Value = data["RunningTimeHour"];
				Sheet.Cells[rowPos, 6].Value = "Hour";

				rowPos = rowPos + 2;
				Sheet.Cells[rowPos, 2].Value = "Output";
				Sheet.Cells[rowPos++, 3].Value = data["Output"];

				Sheet.Cells[rowPos, 2].Value = "Flow";
				Sheet.Cells[rowPos++, 3].Value = data["Flow"];

				Sheet.Cells[rowPos, 2].Value = "PR";
				Sheet.Cells[rowPos++, 3].Value = data["PR"];

				using (var range = Sheet.Cells[rowPos, 2, rowPos, 6])
				{
					range.Style.Font.Bold = true;
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
					range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				}

				var rowFirstTable = rowPos;
				string tableString = data["Table"];
				List<Dictionary<string, string>> table = string.IsNullOrEmpty(tableString) ? new List<Dictionary<string, string>>() : JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(tableString);
				foreach (Dictionary<string, string> dic in table)
				{
					Sheet.Cells[rowPos, 2].Value = dic["Status"];
					Sheet.Cells[rowPos, 3].Value = dic["GroupDowntime"];
					Sheet.Cells[rowPos, 4].Value = dic["Stops"];
					Sheet.Cells[rowPos, 5].Value = dic["Durasi"];
					Sheet.Cells[rowPos++, 6].Value = dic["OEELoss"];
				}

				Sheet.Cells[Sheet.Dimension.Address].AutoFitColumns();
				Sheet.Cells[rowFirstTable, 2, rowPos - 1, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);

				Session["DownloadExcel_FileManager"] = Ep.GetAsByteArray();
				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				SetFalseTempData(ex.Message);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}

		}
		public ActionResult Download()
		{

			if (Session["DownloadExcel_FileManager"] != null)
			{
				byte[] data = Session["DownloadExcel_FileManager"] as byte[];
				return File(data, "application/octet-stream", "ReportLossTree.xlsx");
			}
			else
			{
				return new EmptyResult();
			}
		}
		[HttpPost]
		public ActionResult GetReport(string startDate, string endDate, string area, List<string> category, List<string> flow, List<string> output, List<string> equipment)
		{
			try
			{
				DateTime dtEnd = DateTime.ParseExact(endDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
				string dEnd = dtEnd.ToString("dd-MMM-yy");
                string westEnd = dtEnd.ToString("yyyy-MM-dd");
                dtEnd = dtEnd.AddDays(-1);
				string yesterdayEnd = dtEnd.ToString("dd-MMM-yy");

				DateTime dtStart = DateTime.ParseExact(startDate, "dd-MMM-yy", CultureInfo.InvariantCulture);
				string dStart = dtStart.ToString("yyyy-MM-dd");

                string dataCodeDuration = "", dataGroupDowntime = "", dataDeskripsi = "";

                string connAdoDowntime = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;

                string lineProd = string.Join(",", category.Select(x => "'" + x + "'"));
                if (lineProd != "") lineProd = "AND [LineProd] IN (" + lineProd + ")";

				#region Query : Code Duration

				var myQuery = @"SELECT [PlanUnplan] AS [Kode], SUM([Duration]) AS [Duration], COUNT([PlanUnplan]) AS [Count] 
                    FROM [" + area + @"]
                    WHERE 
	                    [Remark] IN (" + string.Join(",", equipment.Select(x => "'" + x + "'")) + @") AND
                        [StartDateTime] >= '" + dStart + @" 06:00' AND
	                    [StartDateTime] < '" + westEnd + @" 06:00'
                        " + lineProd + @"
                    GROUP BY [PlanUnplan]";

				dataCodeDuration = GetQueryResult(connAdoDowntime, myQuery);

				#endregion

				#region Query : Group Downtime

				var myQuery2 = @"SELECT [Category] AS [GroupDowntime], [Description], SUM([Duration]) AS [Duration], COUNT([Category]) AS [Count] 
                    FROM [" + area + @"]
                    WHERE 
	                    [StartDateTime] >= '" + dStart + @" 06:00' AND
	                    [StartDateTime] < '" + westEnd + @" 06:00' AND
	                    [Remark] IN (" + string.Join(",", equipment.Select(x => "'" + x + "'")) + @") AND
	                    [PlanUnplan] = 'PDT'
                        " + lineProd + @"
                    GROUP BY [Description], [Category]";

				dataGroupDowntime = GetQueryResult(connAdoDowntime, myQuery2);

				#endregion

				#region Query : Deskripsi

				var myQuery3 = @"SELECT [Category] AS [GroupDowntime], [Description], SUM([Duration]) AS [Duration], COUNT([Category]) AS [Count] 
                    FROM [" + area + @"]
                    WHERE 
	                    [StartDateTime] >= '" + dStart + @" 06:00' AND
	                    [StartDateTime] < '" + westEnd + @" 06:00' AND
	                    [Remark] IN (" + string.Join(",", equipment.Select(x => "'" + x + "'")) + @") AND
	                    [PlanUnplan] = 'UPDT'
                        " + lineProd + @"
                    GROUP BY [Description], [Category]";

				dataDeskripsi = GetQueryResult(connAdoDowntime, myQuery3);

				#endregion

				#region Query : PSS 4
				string conn = ConfigurationManager.ConnectionStrings["AdoPSS4Conn2"].ConnectionString;
                string selectFlow = "", selectOutput ="", where = "", pivot = "";
                bool first = true;
                bool firstFlow = true;
                bool firstOutput = true;
                flow.ForEach(f =>
                {
                    if (!first)
                    {
                        pivot += ",";
                        where += ",";
                    }
                    else first = false;
                    pivot += "[" + f + "]";
                    where += "'" + f + "'";

                    if (!firstFlow)
                    {
                        selectFlow += "+";
                    }
                    else firstFlow = false;
                    selectFlow += "[" + f + "]";
                });
                output.ForEach(o =>
                {
                    if (!first)
                    {
                        pivot += ",";
                        where += ",";
                    }
                    else first = false;
                    pivot += "[" + o + "]";
                    where += "'" + o + "'";

                    if (!firstOutput)
                    {
                        selectOutput += "+";
                    }
                    else firstOutput = false;
                    selectOutput += "[" + o + "]";
                });

				var queryPSS4 = @"SELECT 
	                    [BatchIdent], [AverageTob], [TotWetTob], [Hour]
                    FROM ( 
                        SELECT 
		                    [NpssReportBatch].[BatchIdent], 
		                    [Data].[EventDateTimeLocal], 
		                    [Data].[DataName],  
		                    CONVERT(FLOAT, replace([Value],',','.')) AS [Value] 
                        FROM [NpssReportBatch] 
                        RIGHT JOIN ( 
                            SELECT *  
                            FROM [NpssReportBatchData] 
                            WHERE 
                                [EventDateTimeLocal] BETWEEN '" + startDate + " 06:00' AND '" + dEnd + @" 06:00' AND 
                                [DataName] IN (" + where + @") 
                        ) AS [Data] ON [Data].[EventID] = [NpssReportBatch].[BatchID] 
                    ) AS [Table] 
                    PIVOT 
                    ( 
                        SUM([Value]) 
                        FOR [DataName] IN (" + pivot + @") 
                    ) AS [Pivoted]
                    CROSS APPLY (SELECT 
	                    (" + selectFlow + @") AS [AverageTob], 
	                    (" + selectOutput + @") AS [TotWetTob]
                    ) AS CA1([AverageTob], [TotWetTob]) 
                    CROSS APPLY (SELECT ([TotWetTob]/[AverageTob]) AS [Hour]) AS CA2([Hour]) ";

				var dataPSS4 = GetQueryResult(conn, queryPSS4);

				#endregion

				return Json(new { Status = "True", Data1 = dataCodeDuration, Data2 = dataGroupDowntime, Data3 = dataDeskripsi, Data4 = dataPSS4 }, JsonRequestBehavior.AllowGet);

			}
			catch (Exception ex)
			{
				ViewBag.Result = false;
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Error = ex.Message });
			}
		}

		public string GetQueryResult(string connString, string query)
		{
			DataSet dset = new DataSet();
			using (SqlConnection conn = new SqlConnection(connString))
			{
				SqlCommand cmd = new SqlCommand(query, conn);
				SqlDataAdapter da = new SqlDataAdapter(cmd);
				da.Fill(dset);
			}
			var result = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]);
			return result;
		}

		public string DataTableToJSONWithJavaScriptSerializer(DataTable table)
		{
			JavaScriptSerializer jsSerializer = new JavaScriptSerializer();
			List<Dictionary<string, object>> parentRow = new List<Dictionary<string, object>>();
			Dictionary<string, object> childRow;
			foreach (DataRow row in table.Rows)
			{
				childRow = new Dictionary<string, object>();
				foreach (DataColumn col in table.Columns)
				{
					childRow.Add(col.ColumnName, row[col]);
				}
				parentRow.Add(childRow);
			}
			return jsSerializer.Serialize(parentRow);
		}
	}
}
