#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models.LPH;
using Fast.Web.Models;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Fast.Web.Resources;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

#endregion

namespace Fast.Web.Controllers.LPH
{
	[CustomAuthorize("history")]
	public class HistoryController : BaseController<LPHModel>
	{
		#region ::Services::
		private readonly ILPHAppService _lphAppService;
		private readonly ILPHApprovalsAppService _lphApprovalAppService;
		private readonly ILPHComponentsAppService _lphComponentsAppService;
		private readonly ILPHLocationsAppService _lphLocationsAppService;
		private readonly ILPHValuesAppService _lphValuesAppService;
		private readonly ILPHValueHistoriesAppService _lphValueHistoriesAppService;
		private readonly ILPHExtrasAppService _lphExtrasAppService;
		private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
		private readonly ILoggerAppService _logger;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IMenuAppService _menuAppService;
		private readonly IUserAppService _userAppService;

		#endregion

		#region ::Constructor::
		public HistoryController(
		  ILPHAppService lphAppService,
		  ILPHComponentsAppService lphComponentsAppService,
		  ILPHLocationsAppService lphLocationsAppService,
		  ILPHValuesAppService lphValuesAppService,
		  ILPHApprovalsAppService lphApprovalsAppService,
		  ILPHValueHistoriesAppService lphValueHistoriesAppService,
		  ILPHExtrasAppService lphExtrasAppService,
		  ILoggerAppService logger,
		  IReferenceAppService referenceAppService,
		  ILPHSubmissionsAppService lPHSubmissionsAppService,
		  ILocationAppService locationAppService,
		  IEmployeeAppService employeeAppService,
		  IMenuAppService menuAppService,
		  IUserAppService userAppService)
		{
			_lphAppService = lphAppService;
			_lphComponentsAppService = lphComponentsAppService;
			_lphLocationsAppService = lphLocationsAppService;
			_lphValuesAppService = lphValuesAppService;
			_lphApprovalAppService = lphApprovalsAppService;
			_lphValueHistoriesAppService = lphValueHistoriesAppService;
			_lphExtrasAppService = lphExtrasAppService;
			_logger = logger;
			_referenceAppService = referenceAppService;
			_lphSubmissionsAppService = lPHSubmissionsAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_menuAppService = menuAppService;
			_userAppService = userAppService;
		}
		#endregion

		#region ::Public Methods::
		public ActionResult Index()
		{
			LPHModel model = new LPHModel();
			model.Access = GetAccess(WebConstants.MenuSlug.EMPLOYEE, _menuAppService);
			ViewBag.LPHTypeList = BindDropDownLPHType();
			LocationTreeModel LocationTree = GetLocationTreeModel();
			ViewBag.LocationTree = LocationTree;

			ViewBag.Me = AccountID;

			return View(model);
		}

		[HttpPost]
		public ActionResult GetData(string dateFilter = "", string shift = "", string lphtype = "", string source = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			try
			{
				if (dateFilter == "" && shift == "" && lphtype == "" && source == "" && main_source == "" && location1 == "" && location2 == "" && location3 == "")
				{
					dateFilter = Session["History_dateFilter"] == null ? "" : (string)Session["History_dateFilter"];
					shift = Session["History_shift"] == null ? "" : (string)Session["History_shift"];
					lphtype = Session["History_lphtype"] == null ? "" : (string)Session["History_lphtype"];
					source = Session["History_source"] == null ? "" : (string)Session["History_source"];
					main_source = Session["History_main_source"] == null ? "" : (string)Session["History_main_source"];
					location1 = Session["History_location1"] == null ? "" : (string)Session["History_location1"];
					location2 = Session["History_location2"] == null ? "" : (string)Session["History_location2"];
					location3 = Session["History_location3"] == null ? "" : (string)Session["History_location3"];
				}
				else
				{
					Session["History_dateFilter"] = dateFilter;
					Session["History_shift"] = shift;
					Session["History_lphtype"] = lphtype;
					Session["History_source"] = source;
					Session["History_main_source"] = main_source;
					Session["History_location1"] = location1;
					Session["History_location2"] = location2;
					Session["History_location3"] = location3;
				}

				var draw = Request.Form.GetValues("draw").FirstOrDefault();
				var start = Request.Form.GetValues("start").FirstOrDefault();
				var length = Request.Form.GetValues("length").FirstOrDefault();
				var sortColumn = Request.Form.GetValues("columns[" + Request.Form.GetValues("order[0][column]").FirstOrDefault() + "][name]").FirstOrDefault();
				var sortColumnDir = Request.Form.GetValues("order[0][dir]").FirstOrDefault();
				var searchValue = Request.Form.GetValues("search[value]").FirstOrDefault();

				// Paging Size (10,20,50,100)    
				int pageSize = length != null ? Convert.ToInt32(length) : 0;
				int skip = start != null ? Convert.ToInt32(start) : 0;

				// Getting all data submissions   			
				string submissionList = _lphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
				List<LPHSubmissionsModel> submissions = submissionList.DeserializeToLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

				// Modified
				//List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
				//submissions = submissions.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
				//submissions = submissions.Where(x => x.UserID == AccountID).ToList();

				// Getting all data lph               
				string lphList = _lphAppService.GetAll(true);
				List<LPHModel> lphs = lphList.DeserializeToLPHList();

				Dictionary<long, string> locationMap = new Dictionary<long, string>();

				if (!string.IsNullOrEmpty(main_source))
				{
					List<long> locations = new List<long>();

					if (main_source == "MyLocat")
					{
						//langsung anggap aja punya sub
						locations.Add(AccountLocationID);

						string deps = _locationAppService.FindBy("ParentID", AccountLocationID, true);
						var depsM = deps.DeserializeToLocationList();

						foreach (var dep in depsM)
						{
							locations.Add(dep.ID);

							string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
							var subdepsM = subdeps.DeserializeToLocationList();

							foreach (var subdep in subdepsM)
							{
								locations.Add(subdep.ID);
							}
						}

						submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
						//submissions = submissions.Where(x => x.LocationID == AccountLocationID).ToList();
					}
					else if (main_source == "Location")
					{
						if (!string.IsNullOrEmpty(location3))
						{
							submissions = submissions.Where(x => x.LocationID == Int64.Parse(location3)).ToList();
						}
						else if (!string.IsNullOrEmpty(location2))
						{
							locations.Add(Int64.Parse(location2));

							string subdeps = _locationAppService.FindBy("ParentID", location2, true);
							var subdepsM = subdeps.DeserializeToLocationList();

							foreach (var subdep in subdepsM)
							{
								locations.Add(subdep.ID);
							}

							submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
						}
						else if (!string.IsNullOrEmpty(location1))
						{
							locations.Add(Int64.Parse(location1));

							string deps = _locationAppService.FindBy("ParentID", location1, true);
							var depsM = deps.DeserializeToLocationList();

							foreach (var dep in depsM)
							{
								locations.Add(dep.ID);

								string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
								var subdepsM = subdeps.DeserializeToLocationList();

								foreach (var subdep in subdepsM)
								{
									locations.Add(subdep.ID);
								}
							}

							submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
						}
						//kalau kosong semua ya skip
					}
					//jika getall abaikan filtering
				}
				else //location saya
				{
					submissions = submissions.Where(x => x.UserID == AccountID).ToList();
				}

				if (!string.IsNullOrEmpty(dateFilter))
				{
					DateTime dateFL = DateTime.Parse(dateFilter);
					submissions = submissions.Where(x => x.Date == dateFL.Date).ToList();
				}
				if (!string.IsNullOrEmpty(shift))
				{
					submissions = submissions.Where(x => x.Shift.Trim() == shift).ToList();
				}
				if (!string.IsNullOrEmpty(lphtype))
				{
					submissions = submissions.Where(x => x.LPHHeader == lphtype + "Controller").ToList();
				}
				if (!string.IsNullOrEmpty(source))
				{
					if (source == "draft")
						submissions = submissions.Where(x => x.IsDeleted == true).ToList();
					else
						submissions = submissions.Where(x => x.IsDeleted == false).ToList();
				}

				foreach (var item in submissions.ToList())
				{
					var check = lphs.Where(x => x.ID == item.LPHID).FirstOrDefault();
					// chanif: exclude LPH yang sudah dihapus
					if (check == null)
					{
						submissions.Remove(item);
						continue;
					}
					else if (check.IsDeleted)
					{
						submissions.Remove(item);
						continue;
					}

					item.CreatedAt = (DateTime)check.ModifiedDate; // chanif: tidak pernah di-update, jadi bisa jadi acuan datetime create
				}

				foreach (var item in submissions)
				{
					if (item.Machine == null)
						item.Machine = "-";

					if (item.Location == null)
						item.Location = "-";

					item.LPHHeader = item.LPHHeader.Replace("Controller", "");

					string approvals = _lphApprovalAppService.FindBy("LPHSubmissionID", item.ID, true);
					LPHApprovalsModel approvalModel = approvals.DeserializeToLPHApprovalList().LastOrDefault();
					if (approvalModel != null)
					{
						if (item.IsDeleted)
						{
							item.Status = "Draft";
							item.StatusChangedAt = null;
						}
						else if (approvalModel.Status.Trim().ToLower() == "draft")
						{
							//item.Status = "Waiting for Approval";
							item.Status = "Draft";
							item.StatusChangedAt = null;
						}
						else if (approvalModel.Status.Trim() == "")
						{
							//item.Status = "Waiting for Approval";
							item.Status = "Submitted";
							item.StatusChangedAt = (DateTime)approvalModel.ModifiedDate;
						}
						else
						{
							item.Status = approvalModel.Status.Trim();
							item.StatusChangedAt = (DateTime)approvalModel.ModifiedDate;
						}

						if (approvalModel.Status.Trim().ToLower() == "approved")
						{
							item.IsCompleted = true;
						}
					}
				}

				int recordsTotal = submissions.Count();

				if (!string.IsNullOrEmpty(searchValue))
				{
					submissions = submissions.Where(m => (m.Shift != null ? m.Shift.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.LPHHeader != null ? m.LPHHeader.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Machine != null ? m.Machine.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.ModifiedBy != null ? m.ModifiedBy.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Date != null ? m.Date.ToString("dd-MMM-yy").ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.CreatedAt != null ? m.CreatedAt.ToString("dd-MMM-yy").ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Status != null ? m.Status.ToLower().Contains(searchValue.ToLower()) : false)).ToList();

				}

				//submissions = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "id":
								submissions = submissions.OrderBy(x => x.LPHID).ToList();
								break;
							case "shift":
								submissions = submissions.OrderBy(x => x.Shift).ToList();
								break;
							case "date":
								submissions = submissions.OrderBy(x => x.Date).ToList();
								break;
							case "subshift":
								submissions = submissions.OrderBy(x => x.SubShift).ToList();
								break;
							case "lphheader":
								submissions = submissions.OrderBy(x => x.LPHHeader).ToList();
								break;
							case "machine":
								submissions = submissions.OrderBy(x => x.Machine).ToList();
								break;
							case "modifiedby":
								submissions = submissions.OrderBy(x => x.ModifiedBy).ToList();
								break;
							case "status":
								submissions = submissions.OrderBy(x => x.Status).ToList();
								break;
							case "location":
								submissions = submissions.OrderBy(x => x.Location).ToList();
								break;
							case "createdat":
								submissions = submissions.OrderBy(x => x.CreatedAt).ToList();
								break;
							case "statuschangedat":
								submissions = submissions.OrderBy(x => x.StatusChangedAt).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "id":
								submissions = submissions.OrderByDescending(x => x.LPHID).ToList();
								break;
							case "shift":
								submissions = submissions.OrderByDescending(x => x.Shift).ToList();
								break;
							case "date":
								submissions = submissions.OrderByDescending(x => x.Date).ToList();
								break;
							case "subshift":
								submissions = submissions.OrderByDescending(x => x.SubShift).ToList();
								break;
							case "lphheader":
								submissions = submissions.OrderByDescending(x => x.LPHHeader).ToList();
								break;
							case "machine":
								submissions = submissions.OrderByDescending(x => x.Machine).ToList();
								break;
							case "modifiedby":
								submissions = submissions.OrderByDescending(x => x.ModifiedBy).ToList();
								break;
							case "status":
								submissions = submissions.OrderByDescending(x => x.Status).ToList();
								break;
							case "location":
								submissions = submissions.OrderBy(x => x.Location).ToList();
								break;
							case "createdat":
								submissions = submissions.OrderByDescending(x => x.CreatedAt).ToList();
								break;
							case "statuschangedat":
								submissions = submissions.OrderByDescending(x => x.StatusChangedAt).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = submissions.Count();

				// Paging     
				//var data = submissions.OrderByDescending(x => x.ID).Skip(skip).Take(pageSize).ToList();
				var data = submissions.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<LPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		public List<SelectListItem> BindDropDownLPHType()
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();

			string lphs = _lphAppService.GetAll(true);
			List<LPHModel> lphList = lphs.DeserializeToLPHList();

			foreach (var item in lphList)
			{
				item.LPHType = item.MenuTitle.Replace("Controller", "");
			}

			lphList = lphList.GroupBy(x => x.LPHType).Select(x => x.FirstOrDefault()).ToList();
			foreach (var item in lphList)
			{
				_menuList.Add(new SelectListItem
				{
					Text = item.LPHType.ToString(),
					Value = item.LPHType.ToString()
				});
			}
			return _menuList;
		}

		[HttpPost]
		public ActionResult Delete(long lphid)
		{
			try
			{
				string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
				string deleteLPH = "DELETE FROM LPHExtras WHERE LPHID = " + lphid + ";";
				deleteLPH = deleteLPH + "DELETE FROM LPHSubmissions WHERE LPHID = " + lphid + ";";
				deleteLPH = deleteLPH + "DELETE FROM LPHApprovals WHERE LPHSubmissionID = " + lphid + ";";
				deleteLPH = deleteLPH + "DELETE FROM LPHValues WHERE LPHComponentID IN (SELECT DISTINCT ID FROM LPHComponents WHERE LPHID = " + lphid + ");";
				deleteLPH = deleteLPH + "DELETE FROM LPHComponents WHERE LPHID = " + lphid + ";";
				deleteLPH = deleteLPH + "DELETE FROM LPHs WHERE ID = " + lphid + ";";
				deleteLPH = deleteLPH + "DELETE FROM LPHLocations WHERE LPHID = " + lphid + ";";

				using (SqlConnection connection = new SqlConnection(strConString))
				{
					connection.Open();
					SqlTransaction transaction = connection.BeginTransaction();

					try
					{
						SqlCommand command = new SqlCommand(deleteLPH, connection, transaction);
						command.ExecuteNonQuery();
						transaction.Commit();
					}
					catch
					{
						transaction.Rollback();

						return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
					}
				}
				////// delete LPH location
				////string lphLocation = _lphLocationsAppService.FindByNoTracking("LPHID", lphid.ToString());
				////List<LPHLocationsModel> lphLocationList = lphLocation.DeserializeToLPHLocationList();
				////foreach (var item in lphLocationList)
				////{
				////	_lphLocationsAppService.Remove(item.ID);
				////}

				////// delete LPH submission and approvals
				////string submission = _lphSubmissionsAppService.FindByNoTracking("LPHID", lphid.ToString());
				////List<LPHSubmissionsModel> lphSubmissionList = submission.DeserializeToLPHSubmissionsList();
				////foreach (var item in lphSubmissionList)
				////{
				////	string approvals = _lphApprovalAppService.FindByNoTracking("LPHSubmissionID", item.ID.ToString());
				////	List<LPHApprovalsModel> lphApprovalList = approvals.DeserializeToLPHApprovalList();
				////	foreach (var app in lphApprovalList)
				////	{
				////		_lphApprovalAppService.Remove(app.ID);
				////	}

				////	_lphSubmissionsAppService.Remove(item.ID);
				////}

				////// delete LPH Extras
				////string lphExtra = _lphExtrasAppService.FindByNoTracking("LPHID", lphid.ToString());
				////List<LPHExtrasModel> lphExtraList = lphExtra.DeserializeToLPHExtraList();
				////foreach (var item in lphExtraList)
				////{
				////	_lphExtrasAppService.Remove(item.ID);
				////}

				////// delete LPH Component, Values, and Histories
				////string components = _lphComponentsAppService.FindByNoTracking("LPHID", lphid.ToString());
				////List<LPHComponentsModel> lphComponentList = components.DeserializeToLPHComponentList();
				////foreach (var item in lphComponentList)
				////{
				////	string values = _lphValuesAppService.FindByNoTracking("LPHComponentID", item.ID.ToString());
				////	List<LPHValuesModel> lphValueList = values.DeserializeToLPHValueList();
				////	foreach (var value in lphValueList)
				////	{
				////		string valueHistories = _lphValueHistoriesAppService.FindByNoTracking("LPHValuesID", value.ID.ToString());
				////		List<LPHValueHistoriesModel> lphValueHistoryList = valueHistories.DeserializeToLPHValueHistoryList();
				////		foreach (var valueHistory in lphValueHistoryList)
				////		{
				////			_lphValueHistoriesAppService.Remove(valueHistory.ID);
				////		}

				////		_lphValuesAppService.Remove(value.ID);
				////	}

				////	_lphComponentsAppService.Remove(item.ID);
				////}

				////// delete LPH
				////_lphAppService.Remove(lphid);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult GetLPHDetailByLPHID(long lphid)
		{
			try
			{
				string lph = _lphAppService.GetById(lphid);
				LPHModel model = lph.DeserializeToLPH();
				model.LPHType = model.MenuTitle.Replace("Controller", "");

				return Json(new { Status = "True", Result = model }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult ExportExcel(string dateFilter = "", string shift = "", string lphtype = "", string source = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			try
			{
				// Getting all data submissions   			
				List<LPHSubmissionsModel> submissions = GetSubmissionsExcel(dateFilter, shift, lphtype, source, main_source, location1, location2, location3);

				byte[] excelData = ExcelGenerator.ExportLPHHistory(submissions, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=LPH-SP-History.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult ExportPDF(string dateFilter = "", string shift = "", string lphtype = "", string source = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			try
			{
				// Getting all data    			
				List<LPHSubmissionsModel> submissions = GetSubmissions(dateFilter, shift, lphtype, source, main_source, location1, location2, location3);

				System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#99ccff");
				Document pdfDoc = new Document(PageSize.A3.Rotate(), 10, 10, 10, 10);
				PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
				pdfDoc.Open();

				Image image = Image.GetInstance(Server.MapPath("~/Content/theme/images/fast-blue.jpg"));
				image.ScaleAbsolute(193, 38);
				pdfDoc.Add(image);

				BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.EMBEDDED);
				Font font = new Font(bf, 12);
				pdfDoc.Add(new Paragraph(new Chunk("Title          : LPH SP History", font)));
				pdfDoc.Add(new Paragraph(new Chunk("Generated By   : " + AccountName, font)));
				pdfDoc.Add(new Paragraph(new Chunk("Generated Date : " + DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), font)));

				//Horizontal Line
				Paragraph line = new Paragraph(new Chunk(new LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 1)));
				pdfDoc.Add(line);

				//Table
				PdfPTable table = new PdfPTable(10);
				table.WidthPercentage = 100;
				//0=Left, 1=Centre, 2=Right
				table.HorizontalAlignment = 0;
				table.SpacingBefore = 20f;
				table.SpacingAfter = 30f;

				PdfPCell cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				Chunk chunk = new Chunk("LPH ID");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Date");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Shift");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Machine");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("LPH Type");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Created By");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Location");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Status");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Created At");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = new Color(colFromHex);
				chunk = new Chunk("Status Changed at");
				cell.AddElement(chunk);
				table.AddCell(cell);

				foreach (var item in submissions)
				{
					table.AddCell(item.LPHID.ToString());
					table.AddCell(item.Date.ToString("dd-MMM-yy"));
					table.AddCell(item.Shift);
					table.AddCell(item.Machine);
					table.AddCell(item.LPHHeader);
					table.AddCell(item.ModifiedBy);
					table.AddCell(item.Location);
					table.AddCell(item.Status);
					table.AddCell(item.CreatedAt.ToString("dd-MMM-yy HH:mm"));
					table.AddCell(item.StatusChangedAt.HasValue ? item.StatusChangedAt.Value.ToString("dd-MMM-yy HH:mm") : "");
				}

				pdfDoc.Add(table);
				pdfWriter.CloseStream = false;
				pdfDoc.Close();

				Response.Buffer = true;
				Response.ContentType = "application/pdf";
				Response.AddHeader("content-disposition", "attachment;filename=LPH-SP-History.pdf");
				Response.Cache.SetCacheability(HttpCacheability.NoCache);
				Response.Write(pdfDoc);
				Response.End();
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.GenerateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID);
			}

			return RedirectToAction("Index");
		}
        #endregion
        private List<LPHSubmissionsModel> GetSubmissionsExcel(string dateFilter = "", string shift = "", string lphtype = "", string source = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
        {
            string submissionList = _lphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
            List<LPHSubmissionsModel> submissions = submissionList.DeserializeToLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

            if (!string.IsNullOrEmpty(main_source))
            {
                List<long> locations = new List<long>();

                if (main_source == "MyLocat")
                {
                    //langsung anggap aja punya sub
                    locations.Add(AccountLocationID);

                    string deps = _locationAppService.FindBy("ParentID", AccountLocationID, true);
                    var depsM = deps.DeserializeToLocationList();

                    foreach (var dep in depsM)
                    {
                        locations.Add(dep.ID);

                        string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                        var subdepsM = subdeps.DeserializeToLocationList();

                        foreach (var subdep in subdepsM)
                        {
                            locations.Add(subdep.ID);
                        }
                    }

                    submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                }
                else if (main_source == "Location")
                {
                    if (!string.IsNullOrEmpty(location3))
                    {
                        submissions = submissions.Where(x => x.LocationID == Int64.Parse(location3)).ToList();
                    }
                    else if (!string.IsNullOrEmpty(location2))
                    {
                        locations.Add(Int64.Parse(location2));

                        string subdeps = _locationAppService.FindBy("ParentID", location2, true);
                        var subdepsM = subdeps.DeserializeToLocationList();

                        foreach (var subdep in subdepsM)
                        {
                            locations.Add(subdep.ID);
                        }

                        submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                    }
                    else if (!string.IsNullOrEmpty(location1))
                    {
                        locations.Add(Int64.Parse(location1));

                        string deps = _locationAppService.FindBy("ParentID", location1, true);
                        var depsM = deps.DeserializeToLocationList();

                        foreach (var dep in depsM)
                        {
                            locations.Add(dep.ID);

                            string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                            var subdepsM = subdeps.DeserializeToLocationList();

                            foreach (var subdep in subdepsM)
                            {
                                locations.Add(subdep.ID);
                            }
                        }

                        submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
                    }
                    //kalau kosong semua ya skip
                }
                //jika getall abaikan filtering
            }
            else //location saya
            {
                submissions = submissions.Where(x => x.UserID == AccountID).ToList();
            }

            if (!string.IsNullOrEmpty(dateFilter))
            {
                DateTime dateFL = DateTime.Parse(dateFilter);
                submissions = submissions.Where(x => x.Date == dateFL.Date).ToList();
            }
            if (!string.IsNullOrEmpty(shift))
            {
                submissions = submissions.Where(x => x.Shift.Trim() == shift).ToList();
            }
            if (!string.IsNullOrEmpty(lphtype))
            {
                submissions = submissions.Where(x => x.LPHHeader == lphtype + "Controller").ToList();
            }
            if (!string.IsNullOrEmpty(source))
            {
                if (source == "draft")
                    submissions = submissions.Where(x => x.IsDeleted == true).ToList();
                else
                    submissions = submissions.Where(x => x.IsDeleted == false).ToList();
            }
            //isi status
            foreach(var item in submissions)
            {

                item.LPHHeader = item.LPHHeader.Replace("Controller", "");

                string approvals = _lphApprovalAppService.FindBy("LPHSubmissionID", item.ID, true);
                LPHApprovalsModel approvalModel = approvals.DeserializeToLPHApprovalList().LastOrDefault();
                if (approvalModel != null)
                {
                    if (item.IsDeleted)
                    {
                        item.Status = "Draft";
                        item.StatusChangedAt = null;
                    }
                    else if (approvalModel.Status.Trim().ToLower() == "draft")
                    {
                        //item.Status = "Waiting for Approval";
                        item.Status = "Draft";
                        item.StatusChangedAt = null;
                    }
                    else if (approvalModel.Status.Trim() == "")
                    {
                        //item.Status = "Waiting for Approval";
                        item.Status = "Submitted";
                        item.StatusChangedAt = (DateTime)approvalModel.ModifiedDate;
                    }
                    else
                    {
                        item.Status = approvalModel.Status.Trim();
                        item.StatusChangedAt = (DateTime)approvalModel.ModifiedDate;
                    }

                    if (approvalModel.Status.Trim().ToLower() == "approved")
                    {
                        item.IsCompleted = true;
                    }
                }
            }

            return submissions;
        }

        #region ::Private Method::
        private List<LPHSubmissionsModel> GetSubmissions(string dateFilter = "", string shift = "", string lphtype = "", string source = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			string submissionList = _lphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
			List<LPHSubmissionsModel> submissions = submissionList.DeserializeToLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

			// Getting all data lph               
			string lphList = _lphAppService.GetAll(true);
			List<LPHModel> lphs = lphList.DeserializeToLPHList();

			Dictionary<long, string> locationMap = new Dictionary<long, string>();

			if (!string.IsNullOrEmpty(main_source))
			{
				List<long> locations = new List<long>();

				if (main_source == "MyLocat")
				{
					//langsung anggap aja punya sub
					locations.Add(AccountLocationID);

					string deps = _locationAppService.FindBy("ParentID", AccountLocationID, true);
					var depsM = deps.DeserializeToLocationList();

					foreach (var dep in depsM)
					{
						locations.Add(dep.ID);

						string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
						var subdepsM = subdeps.DeserializeToLocationList();

						foreach (var subdep in subdepsM)
						{
							locations.Add(subdep.ID);
						}
					}

					submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
				}
				else if (main_source == "Location")
				{
					if (!string.IsNullOrEmpty(location3))
					{
						submissions = submissions.Where(x => x.LocationID == Int64.Parse(location3)).ToList();
					}
					else if (!string.IsNullOrEmpty(location2))
					{
						locations.Add(Int64.Parse(location2));

						string subdeps = _locationAppService.FindBy("ParentID", location2, true);
						var subdepsM = subdeps.DeserializeToLocationList();

						foreach (var subdep in subdepsM)
						{
							locations.Add(subdep.ID);
						}

						submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
					}
					else if (!string.IsNullOrEmpty(location1))
					{
						locations.Add(Int64.Parse(location1));

						string deps = _locationAppService.FindBy("ParentID", location1, true);
						var depsM = deps.DeserializeToLocationList();

						foreach (var dep in depsM)
						{
							locations.Add(dep.ID);

							string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
							var subdepsM = subdeps.DeserializeToLocationList();

							foreach (var subdep in subdepsM)
							{
								locations.Add(subdep.ID);
							}
						}

						submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();
					}
					//kalau kosong semua ya skip
				}
				//jika getall abaikan filtering
			}
			else //location saya
			{
				submissions = submissions.Where(x => x.UserID == AccountID).ToList();
			}

			if (!string.IsNullOrEmpty(dateFilter))
			{
				DateTime dateFL = DateTime.Parse(dateFilter);
				submissions = submissions.Where(x => x.Date == dateFL.Date).ToList();
			}
			if (!string.IsNullOrEmpty(shift))
			{
				submissions = submissions.Where(x => x.Shift.Trim() == shift).ToList();
			}
			if (!string.IsNullOrEmpty(lphtype))
			{
				submissions = submissions.Where(x => x.LPHHeader == lphtype + "Controller").ToList();
			}
			if (!string.IsNullOrEmpty(source))
			{
				if (source == "draft")
					submissions = submissions.Where(x => x.IsDeleted == true).ToList();
				else
					submissions = submissions.Where(x => x.IsDeleted == false).ToList();
			}

			foreach (var item in submissions.ToList())
			{
				var check = lphs.Where(x => x.ID == item.LPHID).FirstOrDefault();
				// chanif: exclude LPH yang sudah dihapus
				if (check == null)
				{
					submissions.Remove(item);
					continue;
				}
				else if (check.IsDeleted)
				{
					submissions.Remove(item);
					continue;
				}

				item.CreatedAt = (DateTime)check.ModifiedDate; // chanif: tidak pernah di-update, jadi bisa jadi acuan datetime create
			}

			foreach (var item in submissions)
			{
				//ambil machineinfo
				try
				{
					item.Machine = "-";

					List<QueryFilter> vFilter = new List<QueryFilter>();
					vFilter.Add(new QueryFilter("LPHID", item.LPHID));
					vFilter.Add(new QueryFilter("ComponentName", "generalInfo-MachInfo"));

					string components = _lphComponentsAppService.Find(vFilter);
					List<LPHComponentsModel> lphComponentList = components.DeserializeToLPHComponentList();
					foreach (var vitem in lphComponentList)
					{
						string values = _lphValuesAppService.FindByNoTracking("LPHComponentID", vitem.ID.ToString());
						List<LPHValuesModel> lphValueList = values.DeserializeToLPHValueList();
						foreach (var value in lphValueList)
						{
							item.Machine = value.Value == null ? "-" : value.Value.ToString();
						}
					}
				}
				catch (Exception e)

				{
					var lala = e.InnerException;
					var yeye = e;

				}
				string user = _userAppService.GetById(item.UserID);
				var userM = user.DeserializeToUser();

				if (String.IsNullOrEmpty(userM.Location))
				{
					item.Location = "";
				}
				else
				{
					item.Location = userM.Location;
				}

				item.LPHHeader = item.LPHHeader.Replace("Controller", "");

				string approvals = _lphApprovalAppService.FindBy("LPHSubmissionID", item.ID, true);
				LPHApprovalsModel approvalModel = approvals.DeserializeToLPHApprovalList().LastOrDefault();
				if (approvalModel != null)
				{
					if (item.IsDeleted)
					{
						item.Status = "Draft";
						item.StatusChangedAt = null;
					}
					else if (approvalModel.Status.Trim().ToLower() == "draft")
					{
						//item.Status = "Waiting for Approval";
						item.Status = "Draft";
						item.StatusChangedAt = null;
					}
					else if (approvalModel.Status.Trim() == "")
					{
						//item.Status = "Waiting for Approval";
						item.Status = "Submitted";
						item.StatusChangedAt = (DateTime)approvalModel.ModifiedDate;
					}
					else
					{
						item.Status = approvalModel.Status.Trim();
						item.StatusChangedAt = (DateTime)approvalModel.ModifiedDate;
					}

					if (approvalModel.Status.Trim().ToLower() == "approved")
					{
						item.IsCompleted = true;
					}
				}
			}

			return submissions;
		}

		private LPHSubmissionsModel GetLPHSubsmision(long id)
		{
			string lphSubmission = _lphSubmissionsAppService.GetBy("LPHID", id);
			LPHSubmissionsModel submissionsModel = lphSubmission.DeserializeToLPHSubmissions();

			return submissionsModel;
		}

		private LocationTreeModel GetLocationTreeModel()
		{
			LocationTreeModel model = new LocationTreeModel();

			int index = 1;

			// get production center list
			string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
			List<LocationModel> pcList = pcs.DeserializeToLocationList();
			foreach (var item in pcList)
			{
				string pc = _referenceAppService.GetDetailBy("Code", item.Code, true);
				ProductionCenterModel pcModel = pc.DeserializeToProductionCenter(index++, item.ID, item.ParentID);
				model.ProductionCenters.Add(pcModel);
			}

			// get department list
			foreach (var pc in model.ProductionCenters)
			{
				LocationModel currentPC = pcList.Where(x => x.Code == pc.Code).FirstOrDefault();
				if (currentPC != null)
				{
					string departments = _locationAppService.FindBy("ParentID", currentPC.ID, true);
					List<LocationModel> departmentList = departments.DeserializeToLocationList();

					foreach (var d in departmentList)
					{
						string depts = _referenceAppService.GetDetailBy("Code", d.Code, true);
						DepartmentModel deptModel = depts.DeserializeToDepartment(index++, d.ID, d.ParentID);

						string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
						List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

						foreach (var subdeb in subdepartmentList)
						{
							string sds = _referenceAppService.GetDetailBy("Code", subdeb.Code, true);
							deptModel.SubDepartments.Add(sds.DeserializeToSubDepartment(index++, subdeb.ID, subdeb.ParentID));
						}

						pc.Departments.Add(deptModel);
					}
				}
			}

			return model;
		}

		#endregion
	}
}