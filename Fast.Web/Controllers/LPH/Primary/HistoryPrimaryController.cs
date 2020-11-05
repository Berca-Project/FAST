using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Models;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using iTextSharp.text;
using iTextSharp.text.pdf.draw;
using iTextSharp.text.pdf;
using System.Web;
using Fast.Web.Resources;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace Fast.Web.Controllers.LPH.Primary
{
	[CustomAuthorize("HistoryPrimary")]
	public class HistoryPrimaryController : BaseController<PPLPHModel>
	{
		private readonly IPPLPHAppService _ppLphAppService;
		private readonly IPPLPHApprovalsAppService _ppLphApprovalAppService;
		private readonly IPPLPHComponentsAppService _ppLphComponentsAppService;
		private readonly IPPLPHLocationsAppService _ppLphLocationsAppService;
		private readonly IPPLPHValuesAppService _ppLphValuesAppService;
		private readonly IPPLPHValueHistoriesAppService _ppLphValueHistoriesAppService;
		private readonly IPPLPHExtrasAppService _ppLphExtrasAppService;
		private readonly IPPLPHSubmissionsAppService _ppLphSubmissionsAppService;
		private readonly ILoggerAppService _logger;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IMenuAppService _menuAppService;
		private readonly IUserAppService _userAppService;

		public HistoryPrimaryController(
		  IPPLPHAppService ppLphAppService,
		  IPPLPHComponentsAppService ppLphComponentsAppService,
		  IPPLPHLocationsAppService ppLphLocationsAppService,
		  IPPLPHValuesAppService ppLphValuesAppService,
		  IPPLPHApprovalsAppService ppLphApprovalsAppService,
		  IPPLPHValueHistoriesAppService ppLphValueHistoriesAppService,
		  IPPLPHExtrasAppService ppLphExtrasAppService,
		  IPPLPHSubmissionsAppService ppLphSubmissionsAppService,
		  ILoggerAppService logger,
		  IReferenceAppService referenceAppService,
		  ILocationAppService locationAppService,
		  IEmployeeAppService employeeAppService,
		  IMenuAppService menuAppService,
		  IUserAppService userAppService)
		{
			_ppLphAppService = ppLphAppService;
			_ppLphComponentsAppService = ppLphComponentsAppService;
			_ppLphLocationsAppService = ppLphLocationsAppService;
			_ppLphValuesAppService = ppLphValuesAppService;
			_ppLphApprovalAppService = ppLphApprovalsAppService;
			_ppLphValueHistoriesAppService = ppLphValueHistoriesAppService;
			_ppLphExtrasAppService = ppLphExtrasAppService;
			_ppLphSubmissionsAppService = ppLphSubmissionsAppService;
			_logger = logger;
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_menuAppService = menuAppService;
			_userAppService = userAppService;

		}
		// GET: HistoryPrimary
		public ActionResult Index()
		{
			PPLPHModel model = new PPLPHModel();
			model.Access = GetAccess(WebConstants.MenuSlug.EMPLOYEE, _menuAppService);
			ViewBag.LPHTypeList = BindDropDownPPLPHType().OrderBy(x => x.Text).ToList();
			LocationTreeModel LocationTree = GetLocationTreeModel();
			ViewBag.LocationTree = LocationTree;

			ViewBag.Me = AccountID;
			return View(model);
		}

		#region Get Data
		public List<SelectListItem> BindDropDownPPLPHType()
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();

			string lphs = _ppLphAppService.GetAll();
			List<PPLPHModel> lphList = lphs.DeserializeToPPLPHList();

			foreach (var item in lphList)
			{
				item.LPHType = item.MenuTitle.Replace("Controller", " ");
			}

			lphList = lphList.GroupBy(x => x.LPHType).Select(x => x.FirstOrDefault()).ToList();
			foreach (var item in lphList)
			{
				var text = item.LPHType.ToString().Trim();

				if (text == "LPHPrimaryKretekLineAddback")
					text = "Kretek Line - Addback";
				else if (text == "LPHPrimaryDiet")
					text = "Intermediate Line - DIET";
				else if (text == "LPHPrimaryCloveInfeedConditioning")
					text = "Intermediate Line - Clove Feeding & DCCC";
				else if (text == "LPHPrimaryCSFCutDryPacking")
					text = "Intermediate Line - CSF Cut Dry & Packing";
				else if (text == "LPHPrimaryCSFInfeedConditioning")
					text = "Intermediate Line - CSF Feeding & DCCC";
				else if (text == "LPHPrimaryCloveCutDryPacking")
					text = "Intermediate Line - Clove Cut Dry & Packing";
				else if (text == "LPHPrimaryRTC")
					text = "Intermediate Line - RTC";
				else if (text == "LPHPrimaryKitchen")
					text = "Intermediate Line - Casing Kitchen";
				else if (text == "LPHPrimaryWhiteLineOTP")
					text = "White Line OTP - Process Note";
				else if (text == "LPHPrimaryKretekLineFeeding")
					text = "Kretek Line - Feeding KR & RJ";
				else if (text == "LPHPrimaryKretekLineConditioning")
					text = "Kretek Line - DCCC KR & RJ";
				else if (text == "LPHPrimaryKretekLineCuttingDrying")
					text = "Kretek Line - Cut Dry";
				else if (text == "LPHPrimaryKretekLinePacking")
					text = "Kretek Line - Packing";
				else if (text == "LPHPrimaryCresFeedingConditioning")
					text = "Kretek Line - CRES Feeding & DCCC";
				else if (text == "LPHPrimaryCresDryingPacking")
					text = "Kretek Line - CRES Cut Dry & Packing";
				else if (text == "LPHPrimaryWhiteLineFeedingWhite")
					text = "White Line PMID - Feeding White";
				else if (text == "LPHPrimaryWhiteLineDCCC")
					text = "White Line PMID - DCCC";
				else if (text == "LPHPrimaryWhiteLineCuttingFTD")
					text = "White Line PMID - Cutting + FTD";
				else if (text == "LPHPrimaryWhiteLineAddback")
					text = "White Line PMID - Addback";
				else if (text == "LPHPrimaryWhiteLinePackingWhite")
					text = "White Line PMID - Packing White";
				else if (text == "LPHPrimaryWhiteLineFeedingSPM")
					text = "White Line PMID - Feeding SPM";
				else if (text == "LPHPrimaryISWhiteFeeding")
					text = "White Line PMID - Feeding IS White";
				else if (text == "LPHPrimaryISWhiteCutDry")
					text = "White Line PMID - Cut Dry IS White";

				_menuList.Add(new SelectListItem
				{
					Value = item.LPHType.ToString(),
					Text = text
				});
			}
			return _menuList;
		}

		[HttpPost]
		public ActionResult GetData(string dateFilter = "", string shift = "", string lphtype = "", string source = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			try
			{
                if (dateFilter == "" && shift == "" && lphtype == "" && source == "" && main_source == "" && location1 == "" && location2 == "" && location3 == "")
                {
                    dateFilter = Session["HistoryPrimary_dateFilter"] == null ? "" : (string)Session["HistoryPrimary_dateFilter"];
                    shift = Session["HistoryPrimary_shift"] == null ? "" : (string)Session["HistoryPrimary_shift"];
                    lphtype = Session["HistoryPrimary_lphtype"] == null ? "" : (string)Session["HistoryPrimary_lphtype"];
                    source = Session["HistoryPrimary_source"] == null ? "" : (string)Session["HistoryPrimary_source"];
                    main_source = Session["HistoryPrimary_main_source"] == null ? "" : (string)Session["HistoryPrimary_main_source"];
                    location1 = Session["HistoryPrimary_location1"] == null ? "" : (string)Session["HistoryPrimary_location1"];
                    location2 = Session["HistoryPrimary_location2"] == null ? "" : (string)Session["HistoryPrimary_location2"];
                    location3 = Session["HistoryPrimary_location3"] == null ? "" : (string)Session["HistoryPrimary_location3"];
                }
                else
                {
                    Session["HistoryPrimary_dateFilter"] = dateFilter;
                    Session["HistoryPrimary_shift"] = shift;
                    Session["HistoryPrimary_lphtype"] = lphtype;
                    Session["HistoryPrimary_source"] = source;
                    Session["HistoryPrimary_main_source"] = main_source;
                    Session["HistoryPrimary_location1"] = location1;
                    Session["HistoryPrimary_location2"] = location2;
                    Session["HistoryPrimary_location3"] = location3;
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
				string submissionList = _ppLphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
				List<PPLPHSubmissionsModel> submissions = submissionList.DeserializeToPPLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

				// Getting all data lph               
				string lphList = _ppLphAppService.GetAll(true);
				List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();

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
				else
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
					lphtype = lphtype.Trim();

					if (lphtype == "Kretek Line - Addback")
						lphtype = "LPHPrimaryKretekLineAddback";
					else if (lphtype == "Intermediate Line - DIET")
						lphtype = "LPHPrimaryDiet";
					else if (lphtype == "Intermediate Line - Clove Feeding & DCCC")
						lphtype = "LPHPrimaryCloveInfeedConditioning";
					else if (lphtype == "Intermediate Line - CSF Cut Dry & Packing")
						lphtype = "LPHPrimaryCSFCutDryPacking";
					else if (lphtype == "Intermediate Line - CSF Feeding & DCCC")
						lphtype = "LPHPrimaryCSFInfeedConditioning";
					else if (lphtype == "Intermediate Line - Clove Cut Dry & Packing")
						lphtype = "LPHPrimaryCloveCutDryPacking";
					else if (lphtype == "Intermediate Line - RTC")
						lphtype = "LPHPrimaryRTC";
					else if (lphtype == "Intermediate Line - Casing Kitchen")
						lphtype = "LPHPrimaryKitchen";
					else if (lphtype == "White Line OTP - Process Note")
						lphtype = "LPHPrimaryWhiteLineOTP";
					else if (lphtype == "Kretek Line - Feeding KR & RJ")
						lphtype = "LPHPrimaryKretekLineFeeding";
					else if (lphtype == "Kretek Line - DCCC KR & RJ")
						lphtype = "LPHPrimaryKretekLineConditioning";
					else if (lphtype == "Kretek Line - Cut Dry")
						lphtype = "LPHPrimaryKretekLineCuttingDrying";
					else if (lphtype == "Kretek Line - Packing")
						lphtype = "LPHPrimaryKretekLinePacking";
					else if (lphtype == "Kretek Line - CRES Feeding & DCCC")
						lphtype = "LPHPrimaryCresFeedingConditioning";
					else if (lphtype == "Kretek Line - CRES Cut Dry & Packing")
						lphtype = "LPHPrimaryCresDryingPacking";
					else if (lphtype == "White Line PMID - Feeding White")
						lphtype = "LPHPrimaryWhiteLineFeedingWhite";
					else if (lphtype == "White Line PMID - DCCC")
						lphtype = "LPHPrimaryWhiteLineDCCC";
					else if (lphtype == "White Line PMID - Cutting + FTD")
						lphtype = "LPHPrimaryWhiteLineCuttingFTD";
					else if (lphtype == "White Line PMID - Addback")
						lphtype = "LPHPrimaryWhiteLineAddback";
					else if (lphtype == "White Line PMID - Packing White")
						lphtype = "LPHPrimaryWhiteLinePackingWhite";
					else if (lphtype == "White Line PMID - Feeding SPM")
						lphtype = "LPHPrimaryWhiteLineFeedingSPM";
					else if (lphtype == "White Line PMID - Feeding IS White")
						lphtype = "LPHPrimaryISWhiteFeeding";
					else if (lphtype == "White Line PMID - Cut Dry IS White")
						lphtype = "LPHPrimaryISWhiteCutDry";

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
                    if (item.Location == null)
                        item.Location = "-";

                    item.LPHHeader = item.LPHHeader.Replace("Controller", "");

					string approvals = _ppLphApprovalAppService.FindBy("LPHSubmissionID", item.ID, true);
					PPLPHApprovalsModel approvalModel = approvals.DeserializeToPPLPHApprovalList().LastOrDefault();
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
								submissions = submissions.OrderBy(x => x.ID).ToList();
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
							case "lphtype":
								submissions = submissions.OrderBy(x => x.LPHType).ToList();
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
								submissions = submissions.OrderByDescending(x => x.ID).ToList();
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
							case "lphtype":
								submissions = submissions.OrderByDescending(x => x.LPHType).ToList();
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

				return Json(new { data = new List<PPLPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult GetPPLPHDetailByLPHID(long lphid)
		{
			try
			{
				string lph = _ppLphAppService.GetById(lphid);
				PPLPHModel model = lph.DeserializeToPPLPH();
				model.LPHType = model.MenuTitle.Replace("Controller", "");

				return Json(new { Status = "True", Result = model }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}
		private PPLPHSubmissionsModel GetPPLPHSubmission(long id)
		{
			string lphSubmission = _ppLphSubmissionsAppService.GetBy("LPHID", id);
			PPLPHSubmissionsModel submissionsModel = lphSubmission.DeserializeToPPLPHSubmissions();

			return submissionsModel;
		}

		#endregion

		[HttpPost]
		public ActionResult Delete(long lphid)
		{
			try
			{
                string strConString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                string deleteLPH = "DELETE FROM PPLPHExtras WHERE LPHID = " + lphid + ";";
                deleteLPH = deleteLPH + "DELETE FROM PPLPHSubmissions WHERE LPHID = " + lphid + ";";
                deleteLPH = deleteLPH + "DELETE FROM PPLPHApprovals WHERE LPHSubmissionID = " + lphid + ";";
                deleteLPH = deleteLPH + "DELETE FROM PPLPHValues WHERE LPHComponentID IN (SELECT DISTINCT ID FROM PPLPHComponents WHERE LPHID = " + lphid + ");";
                deleteLPH = deleteLPH + "DELETE FROM PPLPHComponents WHERE LPHID = " + lphid + ";";
                deleteLPH = deleteLPH + "DELETE FROM PPLPHs WHERE ID = " + lphid + ";";
                deleteLPH = deleteLPH + "DELETE FROM PPLPHLocations WHERE LPHID = " + lphid + ";";

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
                //////// delete LPH location
                //////string lphLocation = _ppLphLocationsAppService.FindByNoTracking("LPHID", lphid.ToString());
                //////List<PPLPHLocationsModel> lphLocationList = lphLocation.DeserializeToPPLPHLocationsList();
                //////foreach (var item in lphLocationList)
                //////{
                //////	_ppLphLocationsAppService.Remove(item.ID);
                //////}

                //////// delete LPH submission and approvals
                //////string submission = _ppLphSubmissionsAppService.FindByNoTracking("LPHID", lphid.ToString());
                //////List<PPLPHSubmissionsModel> lphSubmissionList = submission.DeserializeToPPLPHSubmissionsList();
                //////foreach (var item in lphSubmissionList)
                //////{
                //////	string approvals = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", item.ID.ToString());
                //////	List<PPLPHApprovalsModel> lphApprovalList = approvals.DeserializeToPPLPHApprovalList();
                //////	foreach (var app in lphApprovalList)
                //////	{
                //////		_ppLphApprovalAppService.Remove(app.ID);
                //////	}

                //////	_ppLphSubmissionsAppService.Remove(item.ID);
                //////}

                //////// delete LPH Extras
                //////string lphExtra = _ppLphExtrasAppService.FindByNoTracking("LPHID", lphid.ToString());
                //////List<PPLPHExtrasModel> lphExtraList = lphExtra.DeserializeToPPLPHExtrasList();
                //////foreach (var item in lphExtraList)
                //////{
                //////	_ppLphExtrasAppService.Remove(item.ID);
                //////}

                //////// delete LPH Component, Values, and Histories
                //////string components = _ppLphComponentsAppService.FindByNoTracking("LPHID", lphid.ToString());
                //////List<PPLPHComponentsModel> lphComponentList = components.DeserializeToPPLPHComponentList();
                //////foreach (var item in lphComponentList)
                //////{
                //////	string values = _ppLphValuesAppService.FindByNoTracking("LPHComponentID", item.ID.ToString());
                //////	List<PPLPHValuesModel> lphValueList = values.DeserializeToPPLPHValueList();
                //////	foreach (var value in lphValueList)
                //////	{
                //////		string valueHistories = _ppLphValueHistoriesAppService.FindByNoTracking("LPHValuesID", value.ID.ToString());
                //////		List<PPLPHValueHistoriesModel> lphValueHistoryList = valueHistories.DeserializeToPPLPHValueHistoryList();
                //////		foreach (var valueHistory in lphValueHistoryList)
                //////		{
                //////			_ppLphValueHistoriesAppService.Remove(valueHistory.ID);
                //////		}

                //////		_ppLphValuesAppService.Remove(value.ID);
                //////	}

                //////	_ppLphComponentsAppService.Remove(item.ID);
                //////}

                //////// delete LPH
                //////_ppLphAppService.Remove(lphid);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
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
				List<PPLPHSubmissionsModel> submissions = GetSubmissions(dateFilter, shift, lphtype, source, main_source, location1, location2, location3);

				byte[] excelData = ExcelGenerator.ExportLPHHistoryPrimary(submissions, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=LPH-PP-History.xlsx");
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
				List<PPLPHSubmissionsModel> submissions = GetSubmissions(dateFilter, shift, lphtype, source, main_source, location1, location2, location3);

				System.Drawing.Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#99ccff");
				Document pdfDoc = new Document(PageSize.A3.Rotate(), 10, 10, 10, 10);
				PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
				pdfDoc.Open();

				Image image = Image.GetInstance(Server.MapPath("~/Content/theme/images/fast-blue.jpg"));
				image.ScaleAbsolute(193, 38);
				pdfDoc.Add(image);

				BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.EMBEDDED);
				Font font = new Font(bf, 12);
				pdfDoc.Add(new Paragraph(new Chunk("Title          : LPH PP History", font)));
				pdfDoc.Add(new Paragraph(new Chunk("Generated By   : " + AccountName, font)));
				pdfDoc.Add(new Paragraph(new Chunk("Generated Date : " + DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), font)));

				//Horizontal Line
				Paragraph line = new Paragraph(new Chunk(new LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 1)));
				pdfDoc.Add(line);

				//Table
				PdfPTable table = new PdfPTable(9);
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
				Response.AddHeader("content-disposition", "attachment;filename=LPH-PP-History.pdf");
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

		private List<PPLPHSubmissionsModel> GetSubmissions(string dateFilter = "", string shift = "", string lphtype = "", string source = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			// Getting all data submissions   			
			string submissionList = _ppLphSubmissionsAppService.GetAll(false);  //chanif: ambil semua, karena 0 = draft
			List<PPLPHSubmissionsModel> submissions = submissionList.DeserializeToPPLPHSubmissionsList().OrderByDescending(x => x.Date).ToList();

			// Getting all data lph               
			string lphList = _ppLphAppService.GetAll(true);
			List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();

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
			else
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
				lphtype = lphtype.Trim();

				if (lphtype == "Kretek Line - Addback")
					lphtype = "LPHPrimaryKretekLineAddback";
				else if (lphtype == "Intermediate Line - DIET")
					lphtype = "LPHPrimaryDiet";
				else if (lphtype == "Intermediate Line - Clove Feeding & DCCC")
					lphtype = "LPHPrimaryCloveInfeedConditioning";
				else if (lphtype == "Intermediate Line - CSF Cut Dry & Packing")
					lphtype = "LPHPrimaryCSFCutDryPacking";
				else if (lphtype == "Intermediate Line - CSF Feeding & DCCC")
					lphtype = "LPHPrimaryCSFInfeedConditioning";
				else if (lphtype == "Intermediate Line - Clove Cut Dry & Packing")
					lphtype = "LPHPrimaryCloveCutDryPacking";
				else if (lphtype == "Intermediate Line - RTC")
					lphtype = "LPHPrimaryRTC";
				else if (lphtype == "Intermediate Line - Casing Kitchen")
					lphtype = "LPHPrimaryKitchen";
				else if (lphtype == "White Line OTP - Process Note")
					lphtype = "LPHPrimaryWhiteLineOTP";
				else if (lphtype == "Kretek Line - Feeding KR & RJ")
					lphtype = "LPHPrimaryKretekLineFeeding";
				else if (lphtype == "Kretek Line - DCCC KR & RJ")
					lphtype = "LPHPrimaryKretekLineConditioning";
				else if (lphtype == "Kretek Line - Cut Dry")
					lphtype = "LPHPrimaryKretekLineCuttingDrying";
				else if (lphtype == "Kretek Line - Packing")
					lphtype = "LPHPrimaryKretekLinePacking";
				else if (lphtype == "Kretek Line - CRES Feeding & DCCC")
					lphtype = "LPHPrimaryCresFeedingConditioning";
				else if (lphtype == "Kretek Line - CRES Cut Dry & Packing")
					lphtype = "LPHPrimaryCresDryingPacking";
				else if (lphtype == "White Line PMID - Feeding White")
					lphtype = "LPHPrimaryWhiteLineFeedingWhite";
				else if (lphtype == "White Line PMID - DCCC")
					lphtype = "LPHPrimaryWhiteLineDCCC";
				else if (lphtype == "White Line PMID - Cutting + FTD")
					lphtype = "LPHPrimaryWhiteLineCuttingFTD";
				else if (lphtype == "White Line PMID - Addback")
					lphtype = "LPHPrimaryWhiteLineAddback";
				else if (lphtype == "White Line PMID - Packing White")
					lphtype = "LPHPrimaryWhiteLinePackingWhite";
				else if (lphtype == "White Line PMID - Feeding SPM")
					lphtype = "LPHPrimaryWhiteLineFeedingSPM";
				else if (lphtype == "White Line PMID - Feeding IS White")
					lphtype = "LPHPrimaryISWhiteFeeding";
				else if (lphtype == "White Line PMID - Cut Dry IS White")
					lphtype = "LPHPrimaryISWhiteCutDry";

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

				string approvals = _ppLphApprovalAppService.FindBy("LPHSubmissionID", item.ID, true);
				PPLPHApprovalsModel approvalModel = approvals.DeserializeToPPLPHApprovalList().LastOrDefault();
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

			return model;
		}
	}
}