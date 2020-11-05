#region ::Namespaces::
using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.Script.Serialization;
#endregion

namespace Fast.Web.Controllers.LPH
{
	[CustomAuthorize("lhpripper")]

	public class RipperController : BaseController<LPHModel>
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
		private readonly IMppAppService _mppAppService;
		private readonly IWppAppService _wppAppService;
		private readonly IWeeksAppService _weeksAppService;
		private readonly IUserAppService _userAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		private readonly IMachineAppService _machineAppService;
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		private readonly IBrandAppService _brandAppService;
		private readonly IJobTitleAppService _jobTitleAppService;
		#endregion

		#region ::Constructor::
		public RipperController(
		ILPHAppService lphAppService,
		ILPHComponentsAppService lphComponentsAppService,
		ILPHLocationsAppService lphLocationsAppService,
		ILPHValuesAppService lphValuesAppService,
		ILPHApprovalsAppService lphApprovalsAppService,
		ILPHValueHistoriesAppService lphValueHistoriesAppService,
		ILPHExtrasAppService lphExtrasAppService,
		ILoggerAppService logger,
		IUserRoleAppService userRoleAppService,
		IReferenceAppService referenceAppService,
		ILPHSubmissionsAppService lPHSubmissionsAppService,
		ILocationAppService locationAppService,
		IEmployeeAppService employeeAppService,
		IWppAppService wppAppService,
		IWeeksAppService weeksAppService,
		IMppAppService mppAppService,
		IUserAppService userAppService,
		IMachineAppService machineAppService,
		IReferenceDetailAppService referenceDetailAppService,
		IBrandAppService brandAppService,
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
			_userAppService = userAppService;
			_machineAppService = machineAppService;
			_referenceDetailAppService = referenceDetailAppService;
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

			LoadLPHHeader(model, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, _referenceAppService, 5);

			List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
			string brand = _brandAppService.GetAll();
			List<BrandModel> brandListAll = brand.DeserializeToBrandList();
			List<BrandModel> brandList = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] != 'F').OrderBy(x => x.Code).ToList();

			List<ReferenceDetailModel> brandModelList = new List<ReferenceDetailModel>();
			brandModelList = brandList.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();

			if (brandList.Count > 0)
			{
				brandModelList = brandModelList.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
			}

			ViewBag.BrandCode = brandModelList;

			//List<ReferenceDetailModel> brandModelListFA = new List<ReferenceDetailModel>();
			//List<BrandModel> brandListFA = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] == 'F').OrderBy(x => x.Code).ToList();
			//brandModelListFA = brandListFA.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();
			//if (brandListFA.Count > 0)
			//{
			//	brandModelListFA = brandModelListFA.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
			//}

			ViewBag.BrandCodeFA = LPHSPHelper.GetBrandList(_brandAppService, locationIdList, model.Brands, false, 0).OrderBy(x => x.Code);

			string referenceStringCRRType = _referenceAppService.GetBy("Name", "CRR Type");
			ReferenceModel referenceModelCRRType = referenceStringCRRType.DeserializeToReference();
			ViewBag.CRRType = null;
			if (referenceModelCRRType != null)
			{
				string referenceDetailStringCRRType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCRRType.ID);
				ViewBag.CRRType = referenceDetailStringCRRType.DeserializeToRefDetailList();
			}

            string referenceStringCFType = _referenceAppService.GetBy("Name", "CF Type");
            ReferenceModel referenceModelCFType = referenceStringCFType.DeserializeToReference();
            ViewBag.CFType = null;
            if (referenceModelCFType != null)
            {
                string referenceDetailStringCFType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCFType.ID);
                ViewBag.CFType = referenceDetailStringCFType.DeserializeToRefDetailList();

            }

            string referenceStringRIOutput = _referenceAppService.GetBy("Name", "RIOutput");
			ReferenceModel referenceModelRIOutput = referenceStringRIOutput.DeserializeToReference();
			ViewBag.RIOutput = null;
			if (referenceModelRIOutput != null)
			{
				string referenceDetailStringRIOutput = _referenceDetailAppService.FindBy("ReferenceID", referenceModelRIOutput.ID);
				ViewBag.RIOutput = referenceDetailStringRIOutput.DeserializeToRefDetailList();
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

				return Json(new { Status = "True", LPHID }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		public ActionResult Edit(long lphid)
		{
			var model = new LPHEditModel();

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			LoadLPHHeader(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, _referenceAppService, 5);

			string lph = _lphAppService.GetById(lphid, true);
			LPHModel lphModel = lph.DeserializeToLPH();
			model.LPH = lphModel;
			if (lphModel != null && lphModel.Header != null)
			{
				string submit = _lphSubmissionsAppService.GetBy("LPHID", lphModel.ID.ToString());
				LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

				string approval = _lphApprovalAppService.FindBy("LPHSubmissionID", submitModel.ID, true);
				List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().OrderBy(x => x.ID).ToList();
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

				List<long> locationIdList = _locationAppService.GetLocIDListByLocType(AccountProdCenterID, "productioncenter");
				string brand = _brandAppService.GetAll();
				List<BrandModel> brandListAll = brand.DeserializeToBrandList();
				List<BrandModel> brandList = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] != 'F').OrderBy(x => x.Code).ToList();

				List<ReferenceDetailModel> brandModelList = new List<ReferenceDetailModel>();
				brandModelList = brandList.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();

				if (brandList.Count > 0)
				{
					brandModelList = brandModelList.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
				}

				ViewBag.BrandCode = brandModelList;

				//List<ReferenceDetailModel> brandModelListFA = new List<ReferenceDetailModel>();
				//List<BrandModel> brandListFA = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] == 'F').OrderBy(x => x.Code).ToList();
				//brandModelListFA = brandListFA.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();
				//if (brandListFA.Count > 0)
				//{
				//	brandModelListFA = brandModelListFA.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
				//}

				ViewBag.BrandCodeFA = LPHSPHelper.GetBrandList(_brandAppService, locationIdList, model2.Brands, false, 0).OrderBy(x => x.Code);

				string referenceStringCRRType = _referenceAppService.GetBy("Name", "CRR Type");
				ReferenceModel referenceModelCRRType = referenceStringCRRType.DeserializeToReference();
				ViewBag.CRRType = null;
				if (referenceModelCRRType != null)
				{
					string referenceDetailStringCRRType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCRRType.ID);
					ViewBag.CRRType = referenceDetailStringCRRType.DeserializeToRefDetailList();
				}

                string referenceStringCFType = _referenceAppService.GetBy("Name", "CF Type");
                ReferenceModel referenceModelCFType = referenceStringCFType.DeserializeToReference();
                ViewBag.CFType = null;
                if (referenceModelCFType != null)
                {
                    string referenceDetailStringCFType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCFType.ID);
                    ViewBag.CFType = referenceDetailStringCFType.DeserializeToRefDetailList();

                }

                string referenceStringRIOutput = _referenceAppService.GetBy("Name", "RIOutput");
				ReferenceModel referenceModelRIOutput = referenceStringRIOutput.DeserializeToReference();
				ViewBag.RIOutput = null;
				if (referenceModelRIOutput != null)
				{
					string referenceDetailStringRIOutput = _referenceDetailAppService.FindBy("ReferenceID", referenceModelRIOutput.ID);
					ViewBag.RIOutput = referenceDetailStringRIOutput.DeserializeToRefDetailList();
				}
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

            LoadLPHHeaderApproval(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, _referenceAppService, 5, model.LPH.LocationID);

			var ProdCenter = AccountProdCenterID;
			var Department = AccountDepartmentID;
			if (model.LPH != null && model.LPH.LocationID != 0)
			{
				string location = _locationAppService.GetById(model.LPH.LocationID);
				LocationModel locationModel = location.DeserializeToLocation();
				if (locationModel.ParentID == 1)
				{
					ProdCenter = locationModel.ID;
				}
				else
				{
					string loc = _locationAppService.GetById(locationModel.ParentID);
					LocationModel locModel = loc.DeserializeToLocation();
					if (locModel.ParentID == 1)
					{
						Department = locationModel.ID;
						ProdCenter = locModel.ID;
					}
					else
					{
						string dep = _locationAppService.GetById(locModel.ID);
						LocationModel depModel = dep.DeserializeToLocation();
						Department = depModel.ID;
						ProdCenter = depModel.ParentID;
					}
				}
			}

			List<long> locationIdList = _locationAppService.GetLocIDListByLocType(ProdCenter, "productioncenter");
			string brand = _brandAppService.GetAll();
			List<BrandModel> brandListAll = brand.DeserializeToBrandList();
			List<BrandModel> brandList = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] != 'F').OrderBy(x => x.Code).ToList();

			List<ReferenceDetailModel> brandModelList = new List<ReferenceDetailModel>();
			brandModelList = brandList.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();

			if (brandList.Count > 0)
			{
				brandModelList = brandModelList.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
			}

			ViewBag.BrandCode = brandModelList;

			//List<ReferenceDetailModel> brandModelListFA = new List<ReferenceDetailModel>();
			//List<BrandModel> brandListFA = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] == 'F').OrderBy(x => x.Code).ToList();
			//brandModelListFA = brandListFA.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();
			//if (brandListFA.Count > 0)
			//{
			//	brandModelListFA = brandModelListFA.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
			//}

			ViewBag.BrandCodeFA = LPHSPHelper.GetBrandList(_brandAppService, locationIdList, model2.Brands, false, 0).OrderBy(x => x.Code);

			string referenceStringCRRType = _referenceAppService.GetBy("Name", "CRR Type");
			ReferenceModel referenceModelCRRType = referenceStringCRRType.DeserializeToReference();
			ViewBag.CRRType = null;
			if (referenceModelCRRType != null)
			{
				string referenceDetailStringCRRType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCRRType.ID);
				ViewBag.CRRType = referenceDetailStringCRRType.DeserializeToRefDetailList();
			}

            string referenceStringCFType = _referenceAppService.GetBy("Name", "CF Type");
            ReferenceModel referenceModelCFType = referenceStringCFType.DeserializeToReference();
            ViewBag.CFType = null;
            if (referenceModelCFType != null)
            {
                string referenceDetailStringCFType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCFType.ID);
                ViewBag.CFType = referenceDetailStringCFType.DeserializeToRefDetailList();

            }

            string referenceStringRIOutput = _referenceAppService.GetBy("Name", "RIOutput");
			ReferenceModel referenceModelRIOutput = referenceStringRIOutput.DeserializeToReference();
			ViewBag.RIOutput = null;
			if (referenceModelRIOutput != null)
			{
				string referenceDetailStringRIOutput = _referenceDetailAppService.FindBy("ReferenceID", referenceModelRIOutput.ID);
				ViewBag.RIOutput = referenceDetailStringRIOutput.DeserializeToRefDetailList();
			}
            refGroup(_referenceAppService, _referenceDetailAppService);
            return View(model);
		}

		[HttpPost]
		public async Task<ActionResult> Approval(long id, string comment, List<LPHExtrasModel> detailExtras, LPHSubmissionsModel detailSubmission, List<LPHComponentsModel> detailComponent, List<LPHValuesModel> detailValue)
		{
			try
			{
				if (detailExtras != null)
				{
					for (int i = 0; i < detailExtras.Count(); i++)
					{
						if (detailExtras[i].HeaderName.Trim() == "inputTable" && detailExtras[i].FieldName.Trim() == "datetime")
						{
							// do something with entryGW.Value or entryGW.Key
							if (detailExtras[i].Value != null)
							{
								string barcode = detailExtras[i].Value.Trim();
								char[] separator = { '_' };
								string[] barcodes = barcode.Split(separator);

								List<QueryFilter> filter = new List<QueryFilter>();
								filter.Add(new QueryFilter("LPHID", barcodes[0]));
								filter.Add(new QueryFilter("Value", barcodes[1]));
								filter.Add(new QueryFilter("HeaderName", "crr"));
								filter.Add(new QueryFilter("IsDeleted", "0"));
								string extraa = _lphExtrasAppService.Find(filter);
								var extraData = extraa.DeserializeToLPHExtraList().FirstOrDefault();

								if (extraData != null && extraData.RowNumber != null)
								{
									List<QueryFilter> filter2 = new List<QueryFilter>();
									filter2.Add(new QueryFilter("LPHID", barcodes[0]));
									filter2.Add(new QueryFilter("HeaderName", "crr"));
									filter2.Add(new QueryFilter("RowNumber", extraData.RowNumber.ToString().Trim()));
									filter2.Add(new QueryFilter("Value", "False"));
									filter2.Add(new QueryFilter("FieldName", "Proceeded"));
									filter2.Add(new QueryFilter("IsDeleted", "0"));

									string flagString = _lphExtrasAppService.FindNoTracking(filter2);
									var flag = flagString.DeserializeToLPHExtraList().FirstOrDefault();

									if (flag != null)
									{
										flag.Value = "True";
										string dataFlag = JsonHelper<LPHExtrasModel>.Serialize(flag);
										_lphExtrasAppService.Update(dataFlag);
									}
								}
							}
						}
					}
				}

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

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
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
					string submissions = _lphSubmissionsAppService.FindBy("LPHHeader", "RipperController", true); //ambil yang sudah submit saja
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

			ViewBag.PageHeader = "LPH Ripper";
			if (type == 1)
				ViewBag.PageHeader = "LPH Ripper <span class='tx-success'>(Previous Submit)</span>";

			LPHHeaderModel model2 = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			string lph = _lphAppService.GetById(lphid, true);
			LPHModel lphModel = lph.DeserializeToLPH();
			model.LPH = lphModel;

			LoadLPHHeader(model2, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService, _referenceAppService, 5, model.LPH.LocationID);

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

			var ProdCenter = AccountProdCenterID;
			var Department = AccountDepartmentID;
			if (model.LPH != null && model.LPH.LocationID != 0)
			{
				string location = _locationAppService.GetById(model.LPH.LocationID);
				LocationModel locationModel = location.DeserializeToLocation();
				if (locationModel.ParentID == 1)
				{
					ProdCenter = locationModel.ID;
				}
				else
				{
					string loc = _locationAppService.GetById(locationModel.ParentID);
					LocationModel locModel = loc.DeserializeToLocation();
					if (locModel.ParentID == 1)
					{
						Department = locationModel.ID;
						ProdCenter = locModel.ID;
					}
					else
					{
						string dep = _locationAppService.GetById(locModel.ID);
						LocationModel depModel = dep.DeserializeToLocation();
						Department = depModel.ID;
						ProdCenter = depModel.ParentID;
					}
				}
			}

			List<long> locationIdList = _locationAppService.GetLocIDListByLocType(ProdCenter, "productioncenter");
			string brand = _brandAppService.GetAll();
			List<BrandModel> brandListAll = brand.DeserializeToBrandList();
			List<BrandModel> brandList = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] != 'F').OrderBy(x => x.Code).ToList();

			List<ReferenceDetailModel> brandModelList = new List<ReferenceDetailModel>();
			brandModelList = brandList.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();

			if (brandList.Count > 0)
			{
				brandModelList = brandModelList.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
			}

			ViewBag.BrandCode = brandModelList;

			//List<ReferenceDetailModel> brandModelListFA = new List<ReferenceDetailModel>();
			//List<BrandModel> brandListFA = brandListAll.Where(x => locationIdList.Any(y => y == x.LocationID) && x.Code[0] == 'F').OrderBy(x => x.Code).ToList();
			//brandModelListFA = brandListFA.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();
			//if (brandListFA.Count > 0)
			//{
			//	brandModelListFA = brandModelListFA.Where(x => x.Code != null && brandList.Any(z => z.Code == x.Code)).OrderBy(x => x.Code).ToList();
			//}

			ViewBag.BrandCodeFA = LPHSPHelper.GetBrandList(_brandAppService, locationIdList, model2.Brands, false, 0).OrderBy(x => x.Code);

			string referenceStringCRRType = _referenceAppService.GetBy("Name", "CRR Type");
			ReferenceModel referenceModelCRRType = referenceStringCRRType.DeserializeToReference();
			ViewBag.CRRType = null;
			if (referenceModelCRRType != null)
			{
				string referenceDetailStringCRRType = _referenceDetailAppService.FindBy("ReferenceID", referenceModelCRRType.ID);
				ViewBag.CRRType = referenceDetailStringCRRType.DeserializeToRefDetailList();
			}

			string referenceStringRIOutput = _referenceAppService.GetBy("Name", "RIOutput");
			ReferenceModel referenceModelRIOutput = referenceStringRIOutput.DeserializeToReference();
			ViewBag.RIOutput = null;
			if (referenceModelRIOutput != null)
			{
				string referenceDetailStringRIOutput = _referenceDetailAppService.FindBy("ReferenceID", referenceModelRIOutput.ID);
				ViewBag.RIOutput = referenceDetailStringRIOutput.DeserializeToRefDetailList();
			}
            refGroup(_referenceAppService, _referenceDetailAppService);

            return View(model);
		}

		[HttpPost]
		public ActionResult GetCRR(int Shift)
		{
			try
			{
				string submission = _lphSubmissionsAppService.FindBy("LPHHeader", "GWGeneralController");
				List<LPHSubmissionsModel> lphSub = JsonConvert.DeserializeObject<List<LPHSubmissionsModel>>(submission);
				string dtNow = DateTime.Now.Date.ToString();

				for (int i = 0; i < lphSub.Count(); i++)
				{
					string dtDB = lphSub[i].ModifiedDate.Value.Date.ToString();

					if (String.Equals(lphSub[i].Shift.Trim(), Shift.ToString()) && String.Equals(dtNow, dtDB))
					{
						string extra = _lphExtrasAppService.FindBy("LPHID", lphSub[i].LPHID.ToString(), true);
						List<LPHExtrasModel> extraList = extra.DeserializeToLPHExtraList();
						for (int j = 0; j < extraList.Count; j++) extraList[j].Value = extraList[j].Value != null ? Uri.EscapeDataString(extraList[j].Value) : "";
						return Json(new { Status = true, extra = extraList }, JsonRequestBehavior.AllowGet);
					}
				}

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult GetBarcode(long lphid, string datetime)
		{
			try
			{
				List<QueryFilter> filter = new List<QueryFilter>();
				filter.Add(new QueryFilter("LPHID", lphid));
				filter.Add(new QueryFilter("Value", datetime));
				filter.Add(new QueryFilter("HeaderName", "crr"));
				filter.Add(new QueryFilter("IsDeleted", "0"));

				string extra = _lphExtrasAppService.Find(filter);
				var extraData = extra.DeserializeToLPHExtraList().FirstOrDefault();

				if (extraData != null && extraData.RowNumber != null)
				{
					List<QueryFilter> filter2 = new List<QueryFilter>();
					filter2.Add(new QueryFilter("LPHID", lphid));
					filter2.Add(new QueryFilter("HeaderName", "crr"));
					filter2.Add(new QueryFilter("RowNumber", extraData.RowNumber.ToString().Trim()));
					filter2.Add(new QueryFilter("IsDeleted", "0"));

					string extra2 = _lphExtrasAppService.Find(filter2);
					List<LPHExtrasModel> extraList = extra2.DeserializeToLPHExtraList().OrderByDescending(x => x.ID).ToList();
					LPHExtrasModel flagChecker = extraList.Where(x => x.FieldName == "Proceeded").FirstOrDefault();
					if (flagChecker == null)
					{
						return Json(new { Status = "False", message = "Barcode not found" }, JsonRequestBehavior.AllowGet);
					}
					else
					{
						if (flagChecker.Value == "False")
						{
							return Json(new { Status = "True", message = extraList }, JsonRequestBehavior.AllowGet);
						}
						else
						{
							return Json(new { Status = "False", message = "The barcode has been processed" }, JsonRequestBehavior.AllowGet);
						}
					}
				}
				else
				{
					return Json(new { Status = "False", message = "Barcode not found" }, JsonRequestBehavior.AllowGet);
				}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
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

		private ReferenceDetailModel GetMachineType(long machineTypeID)
		{
			string machineType = _referenceAppService.GetDetailById(machineTypeID, true);
			ReferenceDetailModel machineTypeModel = machineType.DeserializeToRefDetail();

			return machineTypeModel;
		}

		private string GetMachineTypeCode(long machineTypeID)
		{
			return GetMachineType(machineTypeID).Code;
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
