#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
#endregion

namespace Fast.Web.Controllers.LPH
{
	[CustomAuthorize("lphfilter")]

	public class FilterController : BaseController<LPHModel>
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
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IMppAppService _mppAppService;
		private readonly IWppAppService _wppAppService;
		private readonly IWeeksAppService _weeksAppService;
		private readonly IMachineAppService _machineAppService;
		private readonly IUserAppService _userAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		private readonly IBrandAppService _brandAppService;
		private readonly IJobTitleAppService _jobTitleAppService;
		#endregion

		#region ::Constructor::
		public FilterController(
		ILPHAppService lphAppService,
		IBrandAppService brandAppService,
		ILPHComponentsAppService lphComponentsAppService,
		ILPHLocationsAppService lphLocationsAppService,
		ILPHValuesAppService lphValuesAppService,
		ILPHApprovalsAppService lphApprovalsAppService,
		ILPHValueHistoriesAppService lphValueHistoriesAppService,
		ILPHExtrasAppService lphExtrasAppService,
		ILoggerAppService logger,
		IReferenceDetailAppService referenceDetailAppService,
		IReferenceAppService referenceAppService,
		IUserRoleAppService userRoleAppService,
		ILPHSubmissionsAppService lPHSubmissionsAppService,
		ILocationAppService locationAppService,
		IEmployeeAppService employeeAppService,
		IWppAppService wppAppService,
		IWeeksAppService weeksAppService,
		 IMachineAppService machineAppService,
		IMppAppService mppAppService,
		 IUserAppService userAppService,
		 IJobTitleAppService jobTitleAppService)
		{
			_jobTitleAppService = jobTitleAppService;
			_lphAppService = lphAppService;
			_lphComponentsAppService = lphComponentsAppService;
			_lphLocationsAppService = lphLocationsAppService;
			_lphValuesAppService = lphValuesAppService;
			_lphApprovalAppService = lphApprovalsAppService;
			_lphValueHistoriesAppService = lphValueHistoriesAppService;
			_lphExtrasAppService = lphExtrasAppService;
			_logger = logger;
			_userRoleAppService = userRoleAppService;
			_referenceAppService = referenceAppService;
			_lphSubmissionsAppService = lPHSubmissionsAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_mppAppService = mppAppService;
			_weeksAppService = weeksAppService;
			_wppAppService = wppAppService;
			_machineAppService = machineAppService;
			_referenceDetailAppService = referenceDetailAppService;
			_userAppService = userAppService;
			_brandAppService = brandAppService;
		}
		#endregion

		#region ::Public Methods::
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
			else
			{
				ViewBag.Reset = false;
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

			LoadLPHHeader(model, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 3);

			List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
			ViewBag.Maker = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model.Machines, false, false, "ALL");

            string FilterImage = _referenceAppService.GetBy("Name", "FilterImage");
            ReferenceModel referenceModelFilterImage = FilterImage.DeserializeToReference();
            ViewBag.FilterImage = null;
            if (referenceModelFilterImage != null)
            {
                string referenceDetailStringFilterImage = _referenceDetailAppService.FindBy("ReferenceID", referenceModelFilterImage.ID);
                ViewBag.FilterImage = referenceDetailStringFilterImage.DeserializeToRefDetailList();
            }
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
		public ActionResult Create(List<LPHExtrasModel> detailExtras, LPHSubmissionsModel detailSubmission, List<LPHComponentsModel> detailComponent, List<LPHValuesModel> detailValue)
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
					detailExtras,
					detailSubmission,
					detailComponent,
					detailValue,
					GetType().Name);

                Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID.ToString() + " FilterController Create True", Server.MapPath("~/Uploads/"), AccountEmployeeID);
                return Json(new { Status = "True", LPHID }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + AccountEmployeeID.ToString() + " FilterController Create Failed" + ex.GetAllMessages().ToString(), Server.MapPath("~/Uploads/"), AccountEmployeeID);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Edit(long lphid)
		{
			var model = new LPHEditModel();
			Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController Start Edit LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			LoadLPHHeader(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 3);

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
				//for (int i = 0; i < extraList.Count; i++)
				//	extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
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
					if (valueModel != null)
					{
						var CompoVal = new LPHCompoValModel();
						CompoVal.Component = component;
						CompoVal.Value = valueModel;
						CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";
						model.CompoVal.Add(CompoVal);
					}
				}

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
				ViewBag.Maker = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model2.Machines, false, false, "ALL");

                string FilterImage = _referenceAppService.GetBy("Name", "FilterImage");
                ReferenceModel referenceModelFilterImage = FilterImage.DeserializeToReference();
                ViewBag.FilterImage = null;
                if (referenceModelFilterImage != null)
                {
                    string referenceDetailStringFilterImage = _referenceDetailAppService.FindBy("ReferenceID", referenceModelFilterImage.ID);
                    ViewBag.FilterImage = referenceDetailStringFilterImage.DeserializeToRefDetailList();
                }
                refGroup(_referenceAppService, _referenceDetailAppService);
				//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController End Edit LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
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
				//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController Start Edit SaveData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
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
					isSubmit, Special);

				LPHSubmissionsModel submitModel = _lphSubmissionsAppService.GetByNoTracking("LPHID", id).DeserializeToLPHSubmissions();
				if (submitModel != null && submitModel.Flag != Special && submitModel.Flag != null && Special > 0)
				{
					return Json(new { Status = "False", Msg = "Invalid" }, JsonRequestBehavior.AllowGet);
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
				return Json(new { Status = "False", Msg = "" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult UploadExtrasImages(long id, List<LPHExtrasModel> detailExtras)
		{
			try
			{
				//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController Start Edit Upload Images", Server.MapPath("~/Uploads/"), AccountEmployeeID);

				UploadExtrasImages(id, detailExtras, GetType().Name);

				//Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController End Edit Upload Images True", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController End Edit Upload Images False Catch", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Delete(int id)
		{
			return View();
		}

		[HttpPost]
		public ActionResult DeleteLPH(long id)
		{
			try
			{
				LPHModel lphModel = GetLPH(id);
				lphModel.IsDeleted = true;

				string data = JsonHelper<LPHModel>.Serialize(lphModel);
				_lphAppService.Update(data);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Approval(long lphid)
		{
			Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController Start Approval LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
			string submit = _lphSubmissionsAppService.GetBy("LPHID", lphid.ToString());
			LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

			string approval = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID, true);
			List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().ToList();
			approvalList = approvalList.Where(x => x.UserEmployeeID == AccountEmployeeID).OrderByDescending(x => x.ID).ToList();
			LPHApprovalsModel apprModel = approvalList.FirstOrDefault();

			if (apprModel != null && apprModel.Status.Trim().ToLower() == "approved")
				return RedirectToAction("Submitted", new { lphid = lphid });

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

            var model = LPHSPHelper.SetupLPHApprovalModel(_lphAppService, _lphExtrasAppService, _lphComponentsAppService, _lphValuesAppService, submitModel.ID, lphid);

            LoadLPHHeaderApproval(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 3, model.LPH.LocationID);

			List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
			ViewBag.Maker = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model2.Machines, false, false, "ALL");

            string FilterImage = _referenceAppService.GetBy("Name", "FilterImage");
            ReferenceModel referenceModelFilterImage = FilterImage.DeserializeToReference();
            ViewBag.FilterImage = null;
            if (referenceModelFilterImage != null)
            {
                string referenceDetailStringFilterImage = _referenceDetailAppService.FindBy("ReferenceID", referenceModelFilterImage.ID);
                ViewBag.FilterImage = referenceDetailStringFilterImage.DeserializeToRefDetailList();
            }
            refGroup(_referenceAppService, _referenceDetailAppService);
            Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController End Approval LoadData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
			return View(model);
		}

		[HttpPost]
		public async Task<ActionResult> Approval(long id, string comment, List<LPHExtrasModel> detailExtras, LPHSubmissionsModel detailSubmission, List<LPHComponentsModel> detailComponent, List<LPHValuesModel> detailValue)
		{
			try
			{
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController Start Approval SaveData", Server.MapPath("~/Uploads/"), AccountEmployeeID);
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

				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController End Approval SaveData TRUE", Server.MapPath("~/Uploads/"), AccountEmployeeID);
				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				Helper.LogErrorMessage(DateTime.Now.ToString() + " " + "FilterController End Approval SaveData FALSE Catch", Server.MapPath("~/Uploads/"), AccountEmployeeID);
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
				string lphSubmission = _lphSubmissionsAppService.GetBy("LPHID", lphid);
				LPHSubmissionsModel submissionsModel = lphSubmission.DeserializeToLPHSubmissions();

				string approval = _lphApprovalAppService.GetAll(true);
				List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().OrderBy(x => x.Date).ToList();

				approvalList = approvalList.Where(x => x.LPHSubmissionID == submissionsModel.ID).ToList();
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
					string submissions = _lphSubmissionsAppService.FindBy("LPHHeader", "FilterController", true); //ambil yang sudah submit saja
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

			ViewBag.PageHeader = "LPH Filter";
			if (type == 1)
				ViewBag.PageHeader = "LPH Filter <span class='tx-success'>(Previous Submit)</span>";

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			string lph = _lphAppService.GetById(lphid, true);
			LPHModel lphModel = lph.DeserializeToLPH();
			model.LPH = lphModel;

			LoadLPHHeader(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, null, 3, model.LPH.LocationID);

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
				if (valueModel != null)
				{
					var CompoVal = new LPHCompoValModel();
					CompoVal.Component = component;
					CompoVal.Value = valueModel;
					CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";
					model.CompoVal.Add(CompoVal);
				}
			}

			List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
			ViewBag.Maker = LPHSPHelper.GetMachineList(_machineAppService, locationIdList, model2.Machines, false, false, "ALL");
            refGroup(_referenceAppService, _referenceDetailAppService);
            return View(model);
		}
		#endregion

		[HttpPost]
		public ActionResult GetLastStock(string machine, long LPHID = -1)
		{
			try
			{
				string LPH = _lphSubmissionsAppService.FindBy("LPHHeader", "FilterController", true);
				int DOSValue = 0;
				int PalletValue = 0;
				int ManualValue = 0;
				int TowValue = 0;

				if (string.IsNullOrEmpty(LPH))
				{
					return Json(new { Status = "True", DOS = 0, pallet = 0, manual = 0, tow = 0 }, JsonRequestBehavior.AllowGet);
				}
				else
				{
					bool found = false;
					List<LPHSubmissionsModel> lphModel = JsonConvert.DeserializeObject<List<LPHSubmissionsModel>>(LPH);
					string locationID = AccountLocation;

					lphModel = lphModel.OrderByDescending(x => x.ID).ToList();
					if (lphModel.Count() == 0)
					{
						return Json(new { Status = "True", DOS = 0, pallet = 0, manual = 0, tow = 0 }, JsonRequestBehavior.AllowGet);
					}

					if (LPHID == -1)
					{
						LPHID = Int64.MaxValue;
					}

					lphModel = lphModel.Where(x => x.ID <= LPHID).ToList();

					for (int i = 0; i < lphModel.Count(); i++)
					{
						string locationLPH = _locationAppService.GetLocationFullCode(lphModel[i].LocationID);
						if (locationLPH.Contains(locationID.Substring(0, 5)) && lphModel[i].LPHID != LPHID)
						{
							string checkDataComps = _lphComponentsAppService.FindBy("LPHID", lphModel[i].LPHID.ToString(), true);
							List<LPHComponentsModel> dataComps = checkDataComps.DeserializeToLPHComponentList();

							LPHComponentsModel mach = dataComps.Where(x => x.ComponentName == "GeneralInfo-MachInfo").FirstOrDefault();
							if (mach != null && !string.IsNullOrEmpty(mach.ID.ToString()))
							{
								string machString = _lphValuesAppService.GetBy("LPHComponentID", mach.ID.ToString());
								LPHValuesModel machValue = machString.DeserializeToLPHValue();
								if (string.Equals(machValue.Value, machine))
								{
									LPHComponentsModel DOS = dataComps.Where(x => x.ComponentName == "ProductInformation-LastStockDOS").FirstOrDefault();
									LPHComponentsModel pallet = dataComps.Where(x => x.ComponentName == "ProductInformation-LastStockPallet").FirstOrDefault();
									LPHComponentsModel manual = dataComps.Where(x => x.ComponentName == "ProductInformation-ManualFeeding").FirstOrDefault();
									LPHComponentsModel tow = dataComps.Where(x => x.ComponentName == "Usage-UsTowStockAkhir").FirstOrDefault();

									string dosString = _lphValuesAppService.GetBy("LPHComponentID", DOS.ID.ToString(), true);
									string palletString = _lphValuesAppService.GetBy("LPHComponentID", pallet.ID.ToString(), true);
									string manualString = _lphValuesAppService.GetBy("LPHComponentID", manual.ID.ToString(), true);
									string towString = _lphValuesAppService.GetBy("LPHComponentID", tow.ID.ToString(), true);

									LPHValuesModel DOSVal = dosString.DeserializeToLPHValue();
									LPHValuesModel PalletVal = palletString.DeserializeToLPHValue();
									LPHValuesModel ManualVal = manualString.DeserializeToLPHValue();
									LPHValuesModel TowVal = towString.DeserializeToLPHValue();

									if (DOSVal.Value != null)
									{
										if (!Int32.TryParse(DOSVal.Value, out DOSValue))
										{
											DOSValue = 0;
										}
									}
									if (PalletVal.Value != null)
									{
										if (!Int32.TryParse(PalletVal.Value, out PalletValue))
										{
											PalletValue = 0;
										}
									}
									if (ManualVal.Value != null)
									{
										if (!Int32.TryParse(ManualVal.Value, out ManualValue))
										{
											ManualValue = 0;
										}
									}
									if (TowVal.Value != null)
									{
										if (!Int32.TryParse(TowVal.Value, out TowValue))
										{
											TowValue = 0;
										}
									}
									found = true;
									return Json(new { Status = "True", DOS = DOSValue, pallet = PalletValue, manual = ManualValue, tow = TowValue }, JsonRequestBehavior.AllowGet);
								}
							}
						}
					}

					if (found == false)
					{
						return Json(new { Status = "True", DOS = 0, pallet = 0, manual = 0, tow = 0 }, JsonRequestBehavior.AllowGet);
					}
					else
					{
						return Json(new { Status = "True", DOS = DOSValue, pallet = PalletValue, manual = ManualValue, tow = TowValue }, JsonRequestBehavior.AllowGet);
					}
				}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}

		}

		#region ::Private Methods::
		private string GetFullName(long userID)
		{
			string user = _userAppService.GetById(userID);
			UserModel userModel = user.DeserializeToUser();

			string emp = _employeeAppService.GetBy("EmployeeID", userModel.EmployeeID, true);
			EmployeeModel empModel = emp.DeserializeToEmployee();

			return empModel.FullName;
		}

		private LPHModel GetLPH(long id)
		{
			string data = _lphAppService.GetById(id, true);
			LPHModel lphModel = data.DeserializeToLPH();

			return lphModel;
		}
		#endregion
	}
}
