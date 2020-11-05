#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
#endregion

namespace Fast.Web.Controllers.LPH
{
	[CustomAuthorize("lphapproval")]
	public class ApprovalController : BaseController<LPHApprovalsModel>
	{
		#region ::Services::
		private readonly ILPHAppService _lphAppService;
		private readonly ILPHApprovalsAppService _lphApprovalAppService;
		private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
		private readonly ILoggerAppService _logger;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IUserAppService _userAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILPHComponentsAppService _lphComponentsAppService;
		private readonly ILPHValuesAppService _lphValuesAppService;
		private readonly ILPHExtrasAppService _lphExtrasAppService;
		#endregion

		#region ::Constructor::
		public ApprovalController(
		  ILPHAppService lphAppService,
		  ILPHApprovalsAppService lphApprovalsAppService,
		  ILoggerAppService logger,
		  ILPHSubmissionsAppService lPHSubmissionsAppService,
		  IEmployeeAppService employeeAppService,
		  ILocationAppService locationAppService,
		  IUserAppService userAppService,
		  IReferenceAppService referenceAppService,
		  ILPHComponentsAppService lphComponentsAppService,
		  ILPHValuesAppService lphValuesAppService,
		  ILPHExtrasAppService lphExtrasAppService
		  )
		{
			_logger = logger;
			_lphAppService = lphAppService;
			_lphApprovalAppService = lphApprovalsAppService;
			_lphSubmissionsAppService = lPHSubmissionsAppService;
			_employeeAppService = employeeAppService;
			_userAppService = userAppService;
			_locationAppService = locationAppService;
			_referenceAppService = referenceAppService;
			_lphComponentsAppService = lphComponentsAppService;
			_lphValuesAppService = lphValuesAppService;
			_lphExtrasAppService = lphExtrasAppService;
		}
		#endregion

		#region ::Public Methods::
		public ActionResult Index()
		{
			ViewBag.LPHTypeList = BindDropDownLPHType();
			LocationTreeModel LocationTree = GetLocationTreeModel();
			ViewBag.LocationTree = LocationTree;
			return View();
		}

		[HttpPost]
		public ActionResult GetData(string dateFilter = "", string shift = "", string lphtype = "", string status = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			try
			{
				var draw = Request.Form.GetValues("draw").FirstOrDefault();
				var start = Request.Form.GetValues("start").FirstOrDefault();
				var length = Request.Form.GetValues("length").FirstOrDefault();
				var sortColumn = Request.Form.GetValues("columns[" + Request.Form.GetValues("order[0][column]").FirstOrDefault() + "][name]").FirstOrDefault();
				var sortColumnDir = Request.Form.GetValues("order[0][dir]").FirstOrDefault();
				var searchValue = Request.Form.GetValues("search[value]").FirstOrDefault();

				// Paging Size (10,20,50,100)    
				int pageSize = length != null ? Convert.ToInt32(length) : 0;
				int skip = start != null ? Convert.ToInt32(start) : 0;

				// Getting all data approvals   			
				//string approvals = _lphApprovalAppService.FindBy("LocationID", AccountLocationID.ToString(), true);
				string approvals = _lphApprovalAppService.GetAll();
				List<LPHApprovalsModel> approvalList = approvals.DeserializeToLPHApprovalList().ToList();

				// get user id list 
				//List<long> userIdList = GetUserIDList(AccountEmployeeID);
				long currentUserID = AccountID;
				approvalList = approvalList.OrderByDescending(x => x.ID).ToList();

				// Getting all data lph               
				string lphList = _lphAppService.GetAll(true);
				List<LPHModel> lphs = lphList.DeserializeToLPHList();

				// Getting all data lphsubmission               
				string lphSubList = _lphSubmissionsAppService.GetAll(true);
				List<LPHSubmissionsModel> lphSubs = lphSubList.DeserializeToLPHSubmissionsList();

				//chanif: exclude LPH yg sudah dihapus
				foreach (var item in lphSubs.ToList())
				{
					var check = lphs.Where(x => x.ID == item.LPHID).FirstOrDefault();
					// chanif: exclude LPH yang sudah dihapus
					if (check == null)
					{
						lphSubs.Remove(item);
						continue;
					}
					else if (check.IsDeleted)
					{
						lphSubs.Remove(item);
						continue;
					}
				}

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

						lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
						//submissions = submissions.Where(x => x.LocationID == AccountLocationID).ToList();
					}
					else if (main_source == "Location")
					{
						if (!string.IsNullOrEmpty(location3))
						{
							lphSubs = lphSubs.Where(x => x.LocationID == Int64.Parse(location3)).ToList();
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

							lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
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

							lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
						}
						//kalau kosong semua ya skip
					}
					//jika getall abaikan filtering
				}
				else //location saya
				{
					approvalList = approvalList.Where(x => x.ApproverID == currentUserID).ToList();
					// di atas sudah ada kondisi jika approvernya saya
				}
				if (!string.IsNullOrEmpty(dateFilter))
				{
					DateTime dateFL = DateTime.Parse(dateFilter);
					approvalList = approvalList.Where(x => x.Date == dateFL.Date).ToList();
				}
				if (!string.IsNullOrEmpty(shift))
				{
					lphSubs = lphSubs.Where(x => x.Shift.Trim() == shift).ToList();
				}
				if (!string.IsNullOrEmpty(lphtype))
				{
					lphSubs = lphSubs.Where(x => x.LPHHeader == lphtype + "Controller").ToList();
				}


				List<long> sudah = new List<long>();
				foreach (var item in approvalList.ToList())
				{
					if (sudah.Contains(item.LPHSubmissionID))
					{
						approvalList.Remove(item); // chanif: hanya munculkan status approval terakhir
						continue;
					}

					if (item.Status.Trim().ToLower() == "submitted" || item.Status.Trim() == "")
					{
						item.Status = "Waiting for Approval";
						if (item.ApproverID == currentUserID || item.UserID == currentUserID)
							item.IsNeedMyApproval = true;
					}

					var lphsub = lphSubs.Where(x => x.ID == item.LPHSubmissionID).FirstOrDefault();
					if (lphsub == null)
					{
						approvalList.Remove(item); // chanif: jika tidak ada datanya berarti draft, jangan munculkan approval
						continue;
					}
					else
					{
						item.Location = "-";
						item.Machine = "-";
						item.User = "-";

						if (lphsub.Location != null)
							item.Location = lphsub.Location;
						if (lphsub.Machine != null)
							item.Machine = lphsub.Machine;
						if (lphsub.UserFullName != null)
							item.User = lphsub.UserFullName;
					}


					var check = lphs.Where(x => x.ID == lphsub.LPHID).FirstOrDefault();
					// chanif: exclude LPH yang sudah dihapus
					if (check == null)
					{
						approvalList.Remove(item);
						continue;
					}
					else if (check.IsDeleted)
					{
						approvalList.Remove(item);
						continue;
					}

					var temp = check.MenuTitle;
					item.LPHType = temp.Replace("Controller", "");

					sudah.Add(item.LPHSubmissionID);
				}

				if (!string.IsNullOrEmpty(status))
				{
					if (status.Trim().ToLower() == "approved")
						approvalList = approvalList.Where(x => x.Status.Trim().ToLower() == "approved").ToList();
				}
				else
				{
					approvalList = approvalList.Where(x => x.Status == "Waiting for Approval").ToList();
				}

				int recordsTotal = approvalList.Count();

				// Search    - Correction 231019
				if (!string.IsNullOrEmpty(searchValue))
				{
					approvalList = approvalList.Where(m => (m.Shift != null ? m.Shift.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.LPHType != null ? m.LPHType.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.User != null ? m.User.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Approver != null ? m.Approver.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Date != null ? m.Date.ToString("dd-MMM-yy").ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.LPHSubmissionID != 0 ? m.LPHSubmissionID.ToString().Contains(searchValue.ToLower()) : false)).ToList();

				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "lphsubmissionid":
								approvalList = approvalList.OrderBy(x => x.LPHSubmissionID).ToList();
								break;
							case "location":
								approvalList = approvalList.OrderBy(x => x.Location).ToList();
								break;
							case "date":
								approvalList = approvalList.OrderBy(x => x.Date).ToList();
								break;
							case "lphtype":
								approvalList = approvalList.OrderBy(x => x.LPHType).ToList();
								break;
							case "user":
								approvalList = approvalList.OrderBy(x => x.User).ToList();
								break;
							case "approver":
								approvalList = approvalList.OrderBy(x => x.Approver).ToList();
								break;
							case "status":
								approvalList = approvalList.OrderBy(x => x.Status).ToList();
								break;
							case "notes":
								approvalList = approvalList.OrderBy(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "lphsubmissionid":
								approvalList = approvalList.OrderByDescending(x => x.LPHSubmissionID).ToList();
								break;
							case "location":
								approvalList = approvalList.OrderByDescending(x => x.Location).ToList();
								break;
							case "date":
								approvalList = approvalList.OrderByDescending(x => x.Date).ToList();
								break;
							case "lphtype":
								approvalList = approvalList.OrderByDescending(x => x.LPHType).ToList();
								break;
							case "user":
								approvalList = approvalList.OrderByDescending(x => x.User).ToList();
								break;
							case "approver":
								approvalList = approvalList.OrderByDescending(x => x.Approver).ToList();
								break;
							case "status":
								approvalList = approvalList.OrderByDescending(x => x.Status).ToList();
								break;
							case "notes":
								approvalList = approvalList.OrderByDescending(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = approvalList.Count();

				// Paging     
				var data = approvalList.Skip(skip).Take(pageSize).ToList();

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

		public ActionResult ExportExcel(string dateFilter = "", string shift = "", string lphtype = "", string status = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
		{
			try
			{
				// Getting all data submissions   			

				// Getting all data approvals   			
				//string approvals = _lphApprovalAppService.FindBy("LocationID", AccountLocationID.ToString(), true);
				string approvals = _lphApprovalAppService.GetAll();
				List<LPHApprovalsModel> approvalList = approvals.DeserializeToLPHApprovalList().ToList();

				// get user id list 
				//List<long> userIdList = GetUserIDList(AccountEmployeeID);
				long currentUserID = AccountID;
				approvalList = approvalList.OrderByDescending(x => x.ID).ToList();

				// Getting all data lph               
				string lphList = _lphAppService.GetAll(true);
				List<LPHModel> lphs = lphList.DeserializeToLPHList();

				// Getting all data lphsubmission               
				string lphSubList = _lphSubmissionsAppService.GetAll(true);
				List<LPHSubmissionsModel> lphSubs = lphSubList.DeserializeToLPHSubmissionsList();

				//chanif: exclude LPH yg sudah dihapus
				foreach (var item in lphSubs.ToList())
				{
					var check = lphs.Where(x => x.ID == item.LPHID).FirstOrDefault();
					// chanif: exclude LPH yang sudah dihapus
					if (check == null)
					{
						lphSubs.Remove(item);
						continue;
					}
					else if (check.IsDeleted)
					{
						lphSubs.Remove(item);
						continue;
					}
				}

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

						lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
						//submissions = submissions.Where(x => x.LocationID == AccountLocationID).ToList();
					}
					else if (main_source == "Location")
					{
						if (!string.IsNullOrEmpty(location3))
						{
							lphSubs = lphSubs.Where(x => x.LocationID == Int64.Parse(location3)).ToList();
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

							lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
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

							lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
						}
						//kalau kosong semua ya skip
					}
					//jika getall abaikan filtering
				}
				else //location saya
				{
					approvalList = approvalList.Where(x => x.ApproverID == currentUserID).ToList();
					// di atas sudah ada kondisi jika approvernya saya
				}
				if (!string.IsNullOrEmpty(dateFilter))
				{
					DateTime dateFL = DateTime.Parse(dateFilter);
					approvalList = approvalList.Where(x => x.Date == dateFL.Date).ToList();
				}
				if (!string.IsNullOrEmpty(shift))
				{
					lphSubs = lphSubs.Where(x => x.Shift.Trim() == shift).ToList();
				}
				if (!string.IsNullOrEmpty(lphtype))
				{
					lphSubs = lphSubs.Where(x => x.LPHHeader == lphtype + "Controller").ToList();
				}


				List<long> sudah = new List<long>();
				foreach (var item in approvalList.ToList())
				{
					if (sudah.Contains(item.LPHSubmissionID))
					{
						approvalList.Remove(item); // chanif: hanya munculkan status approval terakhir
						continue;
					}

					if (item.Status.Trim().ToLower() == "submitted" || item.Status.Trim() == "")
					{
						item.Status = "Waiting for Approval";
						if (item.ApproverID == currentUserID || item.UserID == currentUserID)
							item.IsNeedMyApproval = true;
					}

					if (locationMap.ContainsKey(item.LocationID))
					{
						string loc;
						locationMap.TryGetValue(item.LocationID, out loc);
						item.Location = loc;
					}
					else
					{
						item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
						locationMap.Add(item.LocationID, item.Location);
					}

					var lphsub = lphSubs.Where(x => x.ID == item.LPHSubmissionID).FirstOrDefault();
					if (lphsub == null)
					{
						approvalList.Remove(item); // chanif: jika tidak ada datanya berarti draft, jangan munculkan approval
						continue;
					}
					var check = lphs.Where(x => x.ID == lphsub.LPHID).FirstOrDefault();
					// chanif: exclude LPH yang sudah dihapus
					if (check == null)
					{
						approvalList.Remove(item);
						continue;
					}
					else if (check.IsDeleted)
					{
						approvalList.Remove(item);
						continue;
					}

					item.Machine = GetMachine(lphsub); ;

					var temp = check.MenuTitle;
					item.LPHType = temp.Replace("Controller", "");

					//item.Approver = GetFullName(item.ApproverID);
					item.User = GetCreatorBySubmissionID(item.LPHSubmissionID);
					sudah.Add(item.LPHSubmissionID);
				}

				if (!string.IsNullOrEmpty(status))
				{
					if (status.Trim().ToLower() == "approved")
						approvalList = approvalList.Where(x => x.Status.Trim().ToLower() == "approved").ToList();
				}
				else
				{
					approvalList = approvalList.Where(x => x.Status == "Waiting for Approval").ToList();
				}

				byte[] excelData = ExcelGenerator.ExportLPHApproval(approvalList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=LPH-SP-Approval.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		private string GetMachine(LPHSubmissionsModel lphsub)
		{
			string result = "-";
			List<QueryFilter> vFilter = new List<QueryFilter>();
			vFilter.Add(new QueryFilter("LPHID", lphsub.LPHID));
			vFilter.Add(new QueryFilter("ComponentName", "generalInfo-MachInfo"));

			string components = _lphComponentsAppService.Find(vFilter);
			List<LPHComponentsModel> lphComponentList = components.DeserializeToLPHComponentList();
			foreach (var vitem in lphComponentList)
			{
				string values = _lphValuesAppService.FindByNoTracking("LPHComponentID", vitem.ID.ToString());
				List<LPHValuesModel> lphValueList = values.DeserializeToLPHValueList();
				foreach (var value in lphValueList)
				{
					result = value.Value == null ? "-" : value.Value.ToString();
				}
			}

			return result;
		}

		[HttpPost]
		public ActionResult GetLPHDetailBySUBSID(long subsID)
		{
			try
			{
				string subs = _lphSubmissionsAppService.GetById(subsID, true);
				LPHSubmissionsModel subsModel = subs.DeserializeToLPHSubmissions();

				string lph = _lphAppService.GetById(subsModel.LPHID);
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

		[HttpPost]
		public ActionResult ExtractExcel(LPHApprovalsModel model)
		{
			try
			{

				UserModel user = (UserModel)Session["UserLogon"];
				if (!user.LocationID.HasValue)
				{
					SetFalseTempData("Location for the logged user is invalid");
					return RedirectToAction("Index");
				}

				if (model.StartDate > model.EndDate)
				{
					SetFalseTempData("Start Date must be less than End Date");
					return RedirectToAction("Index");
				}
				long locationID = model.SubDepartmentID;

				// Getting all data lph               
				string lphs = _lphAppService.GetAll(true);
				List<LPHModel> lphList = lphs.DeserializeToLPHList();
				lphList = lphList.Where(x => x.LocationID == locationID).ToList();
				lphList = lphList.Where(x => x.ModifiedDate >= model.StartDate && x.ModifiedDate <= model.EndDate).ToList();

				string comps = _lphComponentsAppService.GetAll(true);
				List<LPHComponentsModel> componentList = comps.DeserializeToLPHComponentList();

				string values = _lphValuesAppService.GetAll(true);
				List<LPHValuesModel> valueList = values.DeserializeToLPHValueList();

				string extras = _lphExtrasAppService.GetAll(true);
				List<LPHExtrasModel> extraList = extras.DeserializeToLPHExtraList();

				List<string> compList = new List<string>();
				List<string> valList = new List<string>();
				List<Extra> extList = new List<Extra>();

				string type = "";
				foreach (var item in lphList)
				{
					var temp = item.MenuTitle;
					type = temp.Replace("Controller", "");
					break;
				}
				ArrayList myList = new ArrayList(50);
				if (lphList.Count() > 0)
				{
					foreach (var item in lphList)
					{
						long lphId = item.ID;

						componentList = componentList.Where(x => x.LPHID == lphId).ToList();
						extraList = extraList.Where(x => x.LPHID == lphId).ToList();

						foreach (var ext in extraList)
						{
							Extra x = new Extra();
							x.HeaderName = ext.HeaderName;
							x.FieldName = ext.FieldName;
							x.Value = ext.Value;
							extList.Add(x);
						}

						foreach (var comp in componentList)
						{
							compList.Add(comp.ComponentName);
						}
						foreach (var itemComp in componentList)
						{
							var resValue = valueList.Where(x => x.LPHComponentID == itemComp.ID).ToList().FirstOrDefault();
							valList.Add(resValue.Value);
						}

						myList.Add(extraList);
						myList.Add(compList);
						myList.Add(valList);
					}
				}

				byte[] excelData = ExcelGenerator.RawDataExtract(AccountName, type, model.StartDate.ToString("dd-MMM-yy"), model.EndDate.ToString("dd-MMM-yy"), myList);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Raw Data - " + type + ".xlsx");
				Response.BinaryWrite(excelData);
				Response.End();

			}
			catch (Exception ex)
			{
				SetFalseTempData(ex.Message);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}
			return RedirectToAction("Index");
		}
		#endregion

		#region ::Private Methods::
		private string GetFullName(long userID)
		{
			string user = _userAppService.GetById(userID);
			UserModel userModel = user.DeserializeToUser();

			string emp = _employeeAppService.GetBy("EmployeeID", userModel.EmployeeID, true);
			EmployeeModel empModel = emp.DeserializeToEmployee();

			return empModel.FullName;
		}

		private List<long> GetUserIDList(string employeeID)
		{
			List<long> result = new List<long>();
			string emp = _employeeAppService.GetAll();
			List<EmployeeModel> empList = emp.DeserializeToEmployeeList();

			// get level 1
			List<EmployeeModel> levelOneList = empList.Where(x => x.ReportToID1 != null && x.ReportToID1.Trim() == employeeID.Trim()).ToList();
			foreach (var item in levelOneList)
			{
				result.Add(item.ID);

				// get level 2
				List<EmployeeModel> levelTwoList = empList.Where(x => x.ReportToID1 != null && x.ReportToID1.Trim() == item.EmployeeID.Trim()).ToList();
				foreach (var lvl2 in levelTwoList)
				{
					result.Add(lvl2.ID);
				}
			}

			return result;
		}
		#endregion
		private string GetCreatorBySubmissionID(long submissionID)
		{
			string creator = "";
			string approval = _lphApprovalAppService.FindByNoTracking("LPHSubmissionID", submissionID.ToString(), true);
			List<LPHApprovalsModel> approveModel_list = approval.DeserializeToLPHApprovalList();

			var approveModel = approveModel_list.FirstOrDefault();
			creator = GetFullName(approveModel.UserID);

			return creator;

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

		public class Extra
		{
			public string HeaderName { get; set; }
			public string FieldName { get; set; }
			public string Value { get; set; }
		}

	}
}