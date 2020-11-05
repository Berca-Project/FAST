#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
#endregion

namespace Fast.Web.Controllers.LPH
{
	[CustomAuthorize("lphmaker")]

	public class MakerController : BaseController<LPHModel>
	{
		#region ::Services::
		private readonly ILPHAppService _lphAppService;
		private readonly ILPHComponentsAppService _lphComponentsAppService;
		private readonly ILPHLocationsAppService _lphLocationsAppService;
		private readonly ILPHValuesAppService _lphValuesAppService;
		private readonly ILPHApprovalsAppService _lphApprovalAppService;
		private readonly ILPHValueHistoriesAppService _lphValueHistoriesAppService;
		private readonly ILPHExtrasAppService _lphExtrasAppService;
		private readonly ILoggerAppService _logger;
		private readonly IReferenceAppService _referenceAppService;
		private readonly IMppAppService _mppAppService;
		private readonly IWppAppService _wppAppService;
		private readonly IWeeksAppService _weeksAppService;
		private readonly IUserAppService _userAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
		private readonly IMachineAppService _machineAppService;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IBrandAppService _brandAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IJobTitleAppService _jobTitleAppService;
		#endregion

		#region ::Constructor::
		public MakerController(
			ILPHAppService lphAppService,
			IBrandAppService brandAppService,
			IWppAppService wppAppService,
			IWeeksAppService weeksAppService,
			ILPHComponentsAppService lphComponentsAppService,
			ILPHLocationsAppService lphLocationsAppService,
			ILPHValuesAppService lphValuesAppService,
			ILPHApprovalsAppService lphApprovalsAppService,
			ILPHValueHistoriesAppService lphValueHistoriesAppService,
			ILPHExtrasAppService lphExtrasAppService,
			ILoggerAppService logger,
			IUserRoleAppService userRoleAppService,
			IReferenceAppService referenceAppService,
			IUserAppService userAppService,
			IMppAppService mppAppService,
			IReferenceDetailAppService referenceDetailAppService,
			IMachineAppService machineAppService,
			IEmployeeAppService employeeAppService,
			ILocationAppService locationAppService,
			ILPHSubmissionsAppService lPHSubmissionsAppService,
			IJobTitleAppService jobTitleAppService)
		{
			_jobTitleAppService = jobTitleAppService;
			_lphAppService = lphAppService;
			_locationAppService = locationAppService;
			_lphComponentsAppService = lphComponentsAppService;
			_lphLocationsAppService = lphLocationsAppService;
			_lphValuesAppService = lphValuesAppService;
			_lphApprovalAppService = lphApprovalsAppService;
			_lphValueHistoriesAppService = lphValueHistoriesAppService;
			_lphExtrasAppService = lphExtrasAppService;
			_logger = logger;
			_userRoleAppService = userRoleAppService;
			_referenceAppService = referenceAppService;
			_userAppService = userAppService;
			_mppAppService = mppAppService;
			_referenceDetailAppService = referenceDetailAppService;
			_weeksAppService = weeksAppService;
			_wppAppService = wppAppService;
			_lphSubmissionsAppService = lPHSubmissionsAppService;
			_machineAppService = machineAppService;
			_employeeAppService = employeeAppService;
			_brandAppService = brandAppService;
		}
		#endregion

		#region ::Public Methods::
		[HttpPost]
		public ActionResult GetBeratSpec(string brand)
		{
			try
			{
				List<ReferenceDetailModel> referenceResult = new List<ReferenceDetailModel>();
				string getIDReference = _referenceAppService.GetBy("Name", "Berat Cigarette", true);
				ReferenceModel referenceData = getIDReference.DeserializeToReference();
				if (referenceData != null)
				{
					string resultCloveCon = _referenceDetailAppService.FindBy("ReferenceID", referenceData.ID, true);
					List<ReferenceDetailModel> cloveConList = resultCloveCon.DeserializeToRefDetailList();
					referenceResult = cloveConList.Where(x => x.Code == brand).ToList();
				}

				if (referenceResult.Count > 0)
					return Json(new { Status = "True", Data = referenceResult[0].Description }, JsonRequestBehavior.AllowGet);
				else
					return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Index(bool reset = false)
		{
			ViewBag.Reset = -1;
			ViewBag.OneDay = 0;
			if (!reset)
			{
				LPHSubmissionsModel resultID = null;
				int oneDay = 0;
				LPHSPHelper.IsRedirectToEdit2(_lphAppService, _lphSubmissionsAppService, GetType().Name, AccountName, GetShift(), out resultID, out oneDay);
				//return RedirectToAction("Edit", new { lphid = resultID });
				ViewBag.Reset = resultID;
				ViewBag.OneDay = oneDay;
			}
			if (ViewBag.Reset != null && ViewBag.OneDay == 1)
			{
				return RedirectToAction("Edit", new { lphid = ViewBag.Reset.LPHID });
			}
			ViewBag.AccountEmployeeID = AccountEmployeeID;
			ViewBag.AccountName = AccountName;
			ViewBag.AccountSpvEmployeeID = AccountSpvEmployeeID;
			ViewBag.MachineSelected = "";
			ViewBag.BrandSelected = "";

			List<QueryFilter> filter = new List<QueryFilter>();
			filter.Add(new QueryFilter("EmployeeID", AccountEmployeeID));
			filter.Add(new QueryFilter("Date", DateTime.Now.ToString("yyyy-MM-dd")));
			filter.Add(new QueryFilter("Shift", GetShift()));
			filter.Add(new QueryFilter("IsDeleted", "0"));

			string mpps = _mppAppService.Find(filter);
			var mppData = mpps.DeserializeToMppList().FirstOrDefault();

			if (mppData != null)
			{
				ViewBag.MachineSelected = mppData.EmployeeMachine;

				filter = new List<QueryFilter>();
				filter.Add(new QueryFilter("Maker", mppData.EmployeeMachine));
				filter.Add(new QueryFilter("Date", DateTime.Now.ToString("yyyy-MM-dd")));
				filter.Add(new QueryFilter("IsDeleted", "0"));

				string wpps = _wppAppService.Find(filter);
				var wppData = wpps.DeserializeToWppList().FirstOrDefault();

				if (wppData != null)
					ViewBag.BrandSelected = wppData.Brand;
			}
			LPHHeaderModel model = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			LoadLPHHeader(model, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 1);
            refGroup(_referenceAppService, _referenceDetailAppService);
            return View(model);
		}

		public ActionResult Details(int id)
		{
			return View();
		}

		public ActionResult Create()
		{
			return View();
		}

		[HttpPost]
		public ActionResult Create(List<LPHExtrasModel> detailsExtras, LPHSubmissionsModel detailSubmission, List<LPHValuesModel> detailValue, List<LPHComponentsModel> detailComponent)
		{
			try
			{
				long LPHID = CreateLPH(
					_lphAppService,
					_lphLocationsAppService,
					_lphExtrasAppService,
					_lphSubmissionsAppService,
					_lphApprovalAppService,
					_lphComponentsAppService,
					_lphValuesAppService,
					detailsExtras,
					detailSubmission,
					detailComponent,
					detailValue,
					GetType().Name);

				return Json(new { Status = "True", LPHID }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID.ToString() + " MakerController 210 Inner " + ex.InnerException.ToString(), Server.MapPath("~/Uploads/"));
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID.ToString() + " MakerController 211 GetAllMessage " + ex.GetAllMessages().ToString(), Server.MapPath("~/Uploads/"));
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Edit(long lphid)
		{
			//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController Start Edit LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
			var model = new LPHEditModel();

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			LoadLPHHeader(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 1);

			string lph = _lphAppService.GetById(lphid, true);
			LPHModel lphModel = lph.DeserializeToLPH();
			model.LPH = lphModel;
			if (lphModel != null && lphModel.Header != null)
			{
				string submit = _lphSubmissionsAppService.GetBy("LPHID", lphModel.ID.ToString());
				LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

				string approval = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID, true);
				List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().ToList();
				approvalList = approvalList.Where(x => x.UserEmployeeID == AccountEmployeeID).OrderByDescending(x => x.ID).ToList();
				LPHApprovalsModel apprModel = approvalList.FirstOrDefault();

				if (apprModel != null && apprModel.Status != null && (apprModel.Status.Trim().ToLower() == "submitted" || apprModel.Status.Trim().ToLower() == "approved"))
					return RedirectToAction("Submitted", new { lphid = lphModel.ID });
				if(submitModel.UserID != AccountID)
					return RedirectToAction("Submitted", new { lphid = lphModel.ID });

				string extra = _lphExtrasAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
				List<LPHExtrasModel> extraList = extra.DeserializeToLPHExtraList();
				//for (int i = 0; i < extraList.Count; i++) extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
				model.Extras = extraList.OrderBy(x => x.ID).ToList();

				string compo = _lphComponentsAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
				IEnumerable<LPHComponentsModel> compoList = compo.DeserializeToLPHComponentList();
				compoList = compoList.OrderBy(x => x.ID).ToList();

				model.CompoVal = new List<LPHCompoValModel>();

				string values = _lphValuesAppService.FindBy("SubmissionID", submitModel.ID, true);
				List<LPHValuesModel> valueModelList = values.DeserializeToLPHValueList();

				foreach (var component in compoList)
				{
					var valueModel = valueModelList.Where(x => x.LPHComponentID == component.ID).FirstOrDefault();
					if (valueModel == null)
						valueModel = _lphValuesAppService.GetBy("LPHComponentID", component.ID, true).DeserializeToLPHValue();

					var CompoVal = new LPHCompoValModel();
					CompoVal.Component = component;
					CompoVal.Value = valueModel;
					CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";

					model.CompoVal.Add(CompoVal);
				}

				//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController Start Edit LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
                refGroup(_referenceAppService, _referenceDetailAppService);

				//Generate random number
				Random rnd = new Random();
				int flag = rnd.Next(10000, 90000);
				ViewBag.Special = flag.ToString();
				LoadFlag(lphid, flag);

				return View(model);
			}
			else
			{
				return RedirectToAction("Index", "Home");
			}
		}

		[HttpPost]
		public async Task<ActionResult> Edit(long id, List<LPHExtrasModel> detailExtras, LPHSubmissionsModel detailSubmission, List<LPHComponentsModel> detailComponent, List<LPHValuesModel> detailValue, int isSubmit = 0, int Special = 0)
		{
			try
			{
				//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController Start Edit SaveData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				//if (isSubmit == 1 && !string.IsNullOrEmpty(SupervisorEmail))
				//					await EmailSender.SendEmailLPH(SupervisorEmail, AccountSpvName + "(" + AccountSpvEmployeeID + ")", "Approval LPH Submission");

				await EditLPH(
					_employeeAppService,
					_userAppService,
					_lphSubmissionsAppService,
					_lphExtrasAppService,
					_lphComponentsAppService,
					_lphValuesAppService,
					_lphValueHistoriesAppService,
					_lphApprovalAppService,
					id,
					detailExtras,
					detailValue,
					GetType().Name,
					isSubmit,Special);
				LPHSubmissionsModel submitModel = _lphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToLPHSubmissions();
				if(submitModel != null && submitModel.Flag != Special && submitModel.Flag != null && Special > 0)
                {
					return Json(new { Status = "False", Msg="Invalid" }, JsonRequestBehavior.AllowGet);
				}
				else
                {
					return Json(new { Status = "True", Msg = "" }, JsonRequestBehavior.AllowGet);
				}
			}
			catch (Exception ex)
			{
				//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController End Edit SaveData False Catch", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False", Msg="" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Approval(long lphid)
		{
			Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController Start Approval LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);

			string submit = _lphSubmissionsAppService.GetBy("LPHID", lphid.ToString());
			LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

			string approval = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID, true);
			List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().OrderBy(x => x.ID).ToList();
			approvalList = approvalList.Where(x => x.UserEmployeeID == AccountEmployeeID).OrderByDescending(x => x.ID).ToList();
			LPHApprovalsModel apprModel = approvalList.FirstOrDefault();

			if (apprModel != null && apprModel.Status.Trim().ToLower() == "approved")
				return RedirectToAction("Submitted", new { lphid = lphid });

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

            var model = LPHSPHelper.SetupLPHApprovalModel(_lphAppService, _lphExtrasAppService, _lphComponentsAppService, _lphValuesAppService, submitModel.ID, lphid);

            LoadLPHHeaderApproval(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 1, model.LPH.LocationID);

			Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController End Approval LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
            refGroup(_referenceAppService, _referenceDetailAppService);
            return View(model);
		}

		[HttpPost]
		public async Task<ActionResult> Approval(long id, string comment, List<LPHExtrasModel> detailExtras, LPHSubmissionsModel detailSubmission, List<LPHComponentsModel> detailComponent, List<LPHValuesModel> detailValue)
		{
			try
			{
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController Start Approval SaveData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				await ApproveLPH(
					_employeeAppService,
					_userAppService,
					_lphSubmissionsAppService,
					_lphExtrasAppService,
					_lphComponentsAppService,
					_lphValuesAppService,
					_lphValueHistoriesAppService,
					_lphApprovalAppService,
					id,
					comment,
					detailExtras,
					detailValue,
					GetType().Name);

				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController End Approval SaveData TRUE", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "MakerController End Approval SaveData FALSE Catch", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public async Task<ActionResult> Rejected(long id, string comment)
		{
			try
			{
				await RejectLPH(
					_employeeAppService,
					_userAppService, _lphSubmissionsAppService, _lphApprovalAppService, id, comment);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Delete(int id)
		{
			return View();
		}

		[HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				LPHModel lphModel = GetLPH(id);
				lphModel.IsDeleted = true;

				string lphData = JsonHelper<LPHModel>.Serialize(lphModel);
				//_lphAppService.Update(lphData);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpGet]
		public ActionResult Photo()
		{
			Session["val"] = "";
			return View();
		}

		[HttpPost]
		public ActionResult Photo(string Imagename)
		{
			string sss = Session["val"].ToString();

			ViewBag.pic = "http://localhost:1770/WebImages/" + Session["val"].ToString();

			return View();
		}

		public JsonResult Rebind()
		{
			string path = "http://localhost:1770/WebImages/" + Session["val"].ToString();

			return Json(path, JsonRequestBehavior.AllowGet);
		}

		public ActionResult Capture()
		{
			var stream = Request.InputStream;
			string dump;

			using (var reader = new StreamReader(stream))
			{
				dump = reader.ReadToEnd();

				DateTime nm = DateTime.Now;

				string date = nm.ToString("yyyymmddMMss");

				var path = Server.MapPath("~/WebImages/" + date + "test.jpg");

				System.IO.File.WriteAllBytes(path, String_To_Bytes2(dump));

				ViewData["path"] = date + "test.jpg";

				Session["val"] = date + "test.jpg";
			}

			return View("Index");
		}

		[HttpPost]
		public ActionResult GetApproverList(long lphid)
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
				long currentUserID = AccountID;
				string submit = _lphSubmissionsAppService.GetBy("LPHID", lphid.ToString());
				LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

				string approval = _lphApprovalAppService.GetAll(true);
				List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().OrderBy(x => x.Date).ToList();

				approvalList = approvalList.Where(x => x.LPHSubmissionID == submitModel.ID).ToList();
				for (int ip = 0; ip < approvalList.Count; ip++)
				{
					if (ip % 2 == 0)
					{
						approvalList.ElementAt(ip).User = GetFullName(approvalList.ElementAt(ip).UserID);
						approvalList.ElementAt(ip).Role = "Requestor";
					}
					else
					{
						approvalList.ElementAt(ip).User = GetFullName(approvalList.ElementAt(ip).UserID);
						approvalList.ElementAt(ip).Role = "Approver";
					}
				}
				int recordsTotal = approvalList.Count();
				// Search    - Correction 231019
				if (!string.IsNullOrEmpty(searchValue))
				{
					approvalList = approvalList.Where(m => (m.Shift != null ? m.Shift.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.LPHType != null ? m.LPHType.ToLower().Contains(searchValue.ToLower()) : false)).ToList();

				}
				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "shift":
								approvalList = approvalList.OrderBy(x => x.Shift).ToList();
								break;
							case "location":
								approvalList = approvalList.OrderBy(x => x.Location).ToList();
								break;
							case "lphtype":
								approvalList = approvalList.OrderBy(x => x.LPHType).ToList();
								break;

							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "shift":
								approvalList = approvalList.OrderBy(x => x.Shift).ToList();
								break;
							case "location":
								approvalList = approvalList.OrderBy(x => x.Location).ToList();
								break;
							case "lphtype":
								approvalList = approvalList.OrderBy(x => x.LPHType).ToList();
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
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}

		}

		[HttpPost]
		public ActionResult CheckPrev(long lphid = 0, string machine = "")
		{
			try
			{
				if (lphid != 0)
				{
					string submissions = _lphSubmissionsAppService.FindBy("LPHHeader", "MakerController", true); //ambil yang sudah submit saja
					if (!string.IsNullOrEmpty(submissions))
					{
						List<LPHSubmissionsModel> submisList = submissions.DeserializeToLPHSubmissionsList();
						submisList = submisList.Where(x => x.LPHID < lphid && x.LocationID == AccountLocationID).OrderByDescending(x => x.ID).ToList();

						var prevLPHID = submisList.First().LPHID;
						if (machine.Trim() != "")
						{
							foreach (var submisM in submisList)
							{
								string compo = _lphComponentsAppService.FindBy("LPHID", submisM.LPHID, true);
								IEnumerable<LPHComponentsModel> compoList = compo.DeserializeToLPHComponentList();

								if (compoList.Count() > 0)
								{
									var compoMachine = compoList.Where(x => x.ComponentName.ToLower() == "generalinfo-machinfo").FirstOrDefault();
									if (compoMachine != null && compoMachine.ID != 0)
									{
										string value = _lphValuesAppService.GetBy("LPHComponentID", compoMachine.ID, true);
										LPHValuesModel valueModel = value.DeserializeToLPHValue();

										if (valueModel != null && valueModel.Value == machine)
										{
											prevLPHID = submisM.LPHID;
											break;
										}
									}
								}
							}
						}

						if (submisList.Count() > 0)
							return Json(new { Status = true, LPHID = prevLPHID }, JsonRequestBehavior.AllowGet);
					}
				}

				return Json(new { Status = false }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = false }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Submitted(long lphid, int type = 0)
		{
			var model = new LPHEditModel();

			ViewBag.PageHeader = "LPH Maker";
			if (type == 1)
				ViewBag.PageHeader = "LPH Maker <span class='tx-success'>(Previous Submit)</span>";

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			string lph = _lphAppService.GetById(lphid, true);
			LPHModel lphModel = lph.DeserializeToLPH();
			model.LPH = lphModel;

			LoadLPHHeader(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 1, model.LPH.LocationID);

			string extra = _lphExtrasAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			List<LPHExtrasModel> extraList = extra.DeserializeToLPHExtraList();
			model.Extras = extraList.OrderBy(x => x.ID).ToList();

			string compo = _lphComponentsAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			IEnumerable<LPHComponentsModel> compoList = compo.DeserializeToLPHComponentList();
			compoList = compoList.OrderBy(x => x.ID).ToList();

			model.CompoVal = new List<LPHCompoValModel>();

            //tambahan dari fery (ambil submissionID dari LPHSubmissions.ID)
            string submit = _lphSubmissionsAppService.GetBy("LPHID", lphModel.ID.ToString());
            LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

            string values = _lphValuesAppService.FindBy("SubmissionID", submitModel.ID, true);
            List<LPHValuesModel> valueModelList = values.DeserializeToLPHValueList();

			foreach (var component in compoList)
			{
				var valueModel = valueModelList.Where(x => x.LPHComponentID == component.ID).FirstOrDefault();
				if (valueModel == null)
					valueModel = _lphValuesAppService.GetBy("LPHComponentID", component.ID, true).DeserializeToLPHValue();

				var CompoVal = new LPHCompoValModel();
				CompoVal.Component = component;
				CompoVal.Value = valueModel;
				CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";

				model.CompoVal.Add(CompoVal);
			}
            refGroup(_referenceAppService, _referenceDetailAppService);
            return View(model);
		}

		public ActionResult GetCRR(int Shift, string Date, string Machine)
		{
			try
			{
                //01-May-20
                int vtgl = System.Convert.ToInt32(Date.Substring(5, 2));
                string bulan = Date.Substring(8, 3);
                int vbln = 1;
                switch(bulan.ToLower())
                {
                    case "jan": vbln = 1; break;
                    case "feb": vbln = 2; break;
                    case "mar": vbln = 3; break;
                    case "apr": vbln = 4; break;
                    case "may": vbln = 5; break;
                    case "jun": vbln = 6; break;
                    case "jul": vbln = 7; break;
                    case "aug": vbln = 8; break;
                    case "sep": vbln = 9; break;
                    case "oct": vbln = 10; break;
                    case "nov": vbln = 11; break;
                    case "dec": vbln = 12; break;
                };
                int vthn = 2000 + System.Convert.ToInt32(Date.Substring(12, 2));

                DateTime lphDate = new DateTime(vthn, vbln, vtgl);
//                lphDate = DateTime.ParseExact(lphDate.ToString(), "yyyy-MM-dd", CultureInfo.CurrentCulture);
				string submission = _lphSubmissionsAppService.FindBy("LPHHeader", "GWGeneralController");
				List<LPHSubmissionsModel> lphSub = JsonConvert.DeserializeObject<List<LPHSubmissionsModel>>(submission);
				lphSub = lphSub.Where(item => item.Date == lphDate.Date && item.Shift.Trim() == Shift.ToString().Trim()).ToList();


				List<LPHExtrasModel> extraFinal = new List<LPHExtrasModel>();
				List<LPHExtrasModel> extraTMP = new List<LPHExtrasModel>();

				List<String> crrType = new List<String>();

				for (int i = 0; i < lphSub.Count(); i++)
				{
					string extra = _lphExtrasAppService.FindBy("LPHID", lphSub[i].LPHID.ToString(), true);
					if (extra != "")
					{
						List<LPHExtrasModel> extraList = extra.DeserializeToLPHExtraList();
						extraTMP.AddRange(extraList);
						for (int j = 0; j < extraList.Count; j++)
						{
							if (extraList[j].FieldName == "ItemVal")
							{
								if (!crrType.Contains(extraList[j].Value))
								{
									crrType.Add(extraList[j].Value);
								}
							}
						}
						//for (int j = 0; j < extraList.Count; j++) extraList[j].Value = extraList[j].Value != null ? Uri.EscapeDataString(extraList[j].Value) : "";
						//return Json(new { Status = true, extra = extraList }, JsonRequestBehavior.AllowGet);
					}
				}
				foreach (String type in crrType)
				{
					List<LPHExtrasModel> extraTMP2 = extraTMP.Where(item => item.FieldName == "ItemVal" && item.Value == type).ToList();

					List<LPHExtrasModel> extraTMP3 = new List<LPHExtrasModel>();
					foreach (LPHExtrasModel extra in extraTMP2)
					{
						extraTMP3.AddRange(extraTMP.Where(item => item.LPHID == extra.LPHID && item.RowNumber == extra.RowNumber && item.FieldName == "MachineNumVal" && item.Value == Machine).ToList());
					}
					double sumWeight = 0;
					foreach (LPHExtrasModel extra in extraTMP3)
					{
						sumWeight += extraTMP.Where(item => item.FieldName == "WeightVal" && item.LPHID == extra.LPHID && item.RowNumber == extra.RowNumber).Sum(item => double.Parse(item.Value));
					}

					extraFinal.Add(new LPHExtrasModel
					{
						FieldName = type,
						Value = sumWeight.ToString()
					});
				}

				string volumeValue = "0";
				string submissionPacker = _lphSubmissionsAppService.FindBy("LPHHeader", "PackerController");
				List<LPHSubmissionsModel> lphSub2 = JsonConvert.DeserializeObject<List<LPHSubmissionsModel>>(submissionPacker);
				lphSub2 = lphSub2.Where(item => item.Date == lphDate.Date && item.Shift.Trim() == Shift.ToString()).ToList();
				List<LPHSubmissionsModel> lphSub3 = new List<LPHSubmissionsModel>();
				foreach (LPHSubmissionsModel sMachine in lphSub2)
				{


					string component = _lphComponentsAppService.FindBy("LPHID", sMachine.LPHID);
					List<LPHComponentsModel> lphComponent = JsonConvert.DeserializeObject<List<LPHComponentsModel>>(component);
					LPHComponentsModel lphCompVolume = lphComponent.Where(item => item.ComponentName == "generalInfo-MachInfo").ToList().FirstOrDefault();
					if (lphCompVolume != null)
					{
						string value = _lphValuesAppService.GetBy("LPHComponentID", lphCompVolume.ID, true);
						LPHValuesModel valueModel = value.DeserializeToLPHValue();
						if (valueModel.Value == Machine.Replace('M', 'P'))
						{
							lphSub3.Add(sMachine);
						}
					}
				}
				LPHSubmissionsModel lastThisShift = lphSub3.Where(item => item.Shift != null ? item.Shift.Trim() == Shift.ToString() : false).ToList().OrderByDescending(x => x.ID).FirstOrDefault();
				//item.Date == DateTime.Now.Date &&
				if (lastThisShift != null)
				{
					string component = _lphComponentsAppService.FindBy("LPHID", lastThisShift.LPHID);
					List<LPHComponentsModel> lphComponent = JsonConvert.DeserializeObject<List<LPHComponentsModel>>(component);
					LPHComponentsModel lphCompVolume = lphComponent.Where(item => item.ComponentName == "productionResult-volumekg").ToList().FirstOrDefault();

					string value = _lphValuesAppService.GetBy("LPHComponentID", lphCompVolume.ID, true);
					LPHValuesModel valueModel = value.DeserializeToLPHValue();
					volumeValue = valueModel.Value;
				}

				return Json(new { Status = true, extra = extraFinal, volume = volumeValue }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		#endregion

		#region ::Private Methods::
		private byte[] String_To_Bytes2(string strInput)
		{
			int numBytes = (strInput.Length) / 2;

			byte[] bytes = new byte[numBytes];

			for (int x = 0; x < numBytes; ++x)
			{
				bytes[x] = Convert.ToByte(strInput.Substring(x * 2, 2), 16);
			}

			return bytes;
		}

		private string GetFullName(long userID)
		{
			string user = _userAppService.GetById(userID);
			UserModel userModel = user.DeserializeToUser();

			string emp = _employeeAppService.GetBy("EmployeeID", userModel.EmployeeID, true);
			EmployeeModel empModel = emp.DeserializeToEmployee();

			return empModel.FullName;
		}

		private LPHModel GetLPH(long lphID)
		{
			//string lph = _lphAppService.GetById(lphID, true);
			//LPHModel model = lph.DeserializeToLPH();

			//return model;
			return null;
		}
		public ActionResult GetCTW(string brand)
		{
			try
			{
				string codebrand = _brandAppService.GetBy("Code", brand, false);
				BrandModel brandModel = codebrand.DeserializeToBrand();
				Nullable<double> berat = 0;
				if (brandModel != null)
				{
					if (brandModel.CTW != null)
						berat = brandModel.CTW;
					else
						berat = 0;
				}
				return Json(new { Status = "True", Data = berat }, JsonRequestBehavior.AllowGet);

			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}
		#endregion
	}
}
