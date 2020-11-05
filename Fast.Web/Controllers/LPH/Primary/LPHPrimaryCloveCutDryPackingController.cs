using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace Fast.Web.Controllers.LPH
{
	[CustomAuthorize("primaryclovecutdrypacking")]
	public class LPHPrimaryCloveCutDryPackingController : BaseController<PPLPHModel>
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
		private readonly IMppAppService _mppAppService;
		private readonly IWppAppService _wppAppService;
		private readonly IWeeksAppService _weeksAppService;
		private readonly IUserAppService _userAppService;
		private readonly IUserRoleAppService _userRoleAppService;
		private readonly IMachineAppService _machineAppService;
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		private readonly IJobTitleAppService _jobTitleAppService;

		public LPHPrimaryCloveCutDryPackingController(
		IPPLPHAppService ppLPHAppService,
		IPPLPHComponentsAppService ppLPHComponentsAppService,
		IPPLPHLocationsAppService ppLPHLocationsAppService,
		IPPLPHValuesAppService ppLPHValuesAppService,
		IPPLPHApprovalsAppService ppLPHApprovalsAppService,
		IPPLPHValueHistoriesAppService ppLPHValueHistoriesAppService,
		IPPLPHExtrasAppService ppLPHExtrasAppService,
		IPPLPHSubmissionsAppService ppLPHSubmissionsAppService,
		ILoggerAppService logger,
		IReferenceAppService referenceAppService,
		ILocationAppService locationAppService,
		IEmployeeAppService employeeAppService,
		IWppAppService wppAppService,
		IWeeksAppService weeksAppService,
		IMppAppService mppAppService,
		IUserAppService userAppService,
		IUserRoleAppService userRoleAppService,
		IMachineAppService machineAppService,
		IReferenceDetailAppService referenceDetailAppService,
		IJobTitleAppService jobTitleAppService)
		{
			_jobTitleAppService = jobTitleAppService;
			_ppLphAppService = ppLPHAppService;
			_ppLphComponentsAppService = ppLPHComponentsAppService;
			_ppLphLocationsAppService = ppLPHLocationsAppService;
			_ppLphValuesAppService = ppLPHValuesAppService;
			_ppLphApprovalAppService = ppLPHApprovalsAppService;
			_ppLphValueHistoriesAppService = ppLPHValueHistoriesAppService;
			_ppLphExtrasAppService = ppLPHExtrasAppService;
			_ppLphSubmissionsAppService = ppLPHSubmissionsAppService;
			_logger = logger;
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_employeeAppService = employeeAppService;
			_mppAppService = mppAppService;
			_weeksAppService = weeksAppService;
			_wppAppService = wppAppService;
			_userAppService = userAppService;
			_userRoleAppService = userRoleAppService;
			_machineAppService = machineAppService;
			_referenceDetailAppService = referenceDetailAppService;
		}
		// GET: LPHPrimaryCloveCutDryPackingController
		private void GetReferenceTarget()
		{
			//Ambil Reference untuk Target
			ViewBag.Target1 = "";
			ViewBag.Target2 = "";
			ViewBag.Target3 = "";
			ViewBag.Target4 = "";
			ViewBag.Target5 = "";
			ViewBag.Target6 = "";
			ViewBag.Target7 = "";
			ViewBag.Target8 = "";
			ViewBag.Target9 = "";

			string getIDReference = _referenceAppService.FindBy("Name", "TargetCD", true);
            if (AccountProdCenterID == 4 || AccountProdCenterID == 5)
            {
                getIDReference = _referenceAppService.FindBy("Name", "WestTargetCD", true);

            }
            List<ReferenceModel> referenceList = getIDReference.DeserializeToReferenceList();
			if (referenceList.Count > 0)
			{
				string resultCloveCon = _referenceDetailAppService.FindBy("ReferenceID", referenceList[0].ID, true);
				List<ReferenceDetailModel> cloveConList = resultCloveCon.DeserializeToRefDetailList();
				if (cloveConList != null)
					foreach (var value in cloveConList.ToList())
					{
						switch (value.Code)
						{
							case "T1":
								ViewBag.Target1 = value.Description;
								break;
							case "T2":
								ViewBag.Target2 = value.Description;
								break;
							case "T3":
								ViewBag.Target3 = value.Description;
								break;
							case "T4":
								ViewBag.Target4 = value.Description;
								break;
							case "T5":
								ViewBag.Target5 = value.Description;
								break;
							case "T6":
								ViewBag.Target6 = value.Description;
								break;
							case "T7":
								ViewBag.Target7 = value.Description;
								break;
							case "T8":
								ViewBag.Target8 = value.Description;
								break;
							case "T9":
								ViewBag.Target9 = value.Description;
								break;
						}
					}
			}
			// End Reference Target
		}
		public ActionResult Index(bool reset = false, bool restore = false)
		{
			ViewBag.AccountEmployeeID = AccountEmployeeID;
			ViewBag.AccountSpvEmployeeID = AccountSpvEmployeeID;

			GetReferenceTarget();
			// chanif: rubah model biar lebih fleksibel + tambah viewtipe biar mudah bedain
			var model = new PPLPHEditModel();
			model.ViewType = "index";
			model.Header = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			LoadLPHHeaderPP(_locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _jobTitleAppService);

			//chanif: checkprev, jika ada langsung restore
			PPLPHModel result = new PPLPHModel();
			if (!reset)
			{
				List<QueryFilter> lphFilter = new List<QueryFilter>();
				lphFilter.Add(new QueryFilter("MenuTitle", "LPHPrimaryCloveCutDryPackingController"));
				lphFilter.Add(new QueryFilter("ModifiedBy", AccountName));
				lphFilter.Add(new QueryFilter("IsDeleted", "0"));

				string checkDatas = _ppLphAppService.Find(lphFilter);
				result = checkDatas.DeserializeToPPLPHList().OrderBy(x => x.ID).LastOrDefault();

				if (result != null)
				{
					string checkSubmit = _ppLphSubmissionsAppService.GetBy("LPHID", result.ID, false);
					PPLPHSubmissionsModel submitModel = checkSubmit.DeserializeToPPLPHSubmissions();

					if (submitModel.IsDeleted == true)
						restore = true;
				}
			}

			if (restore)
			{
				#region Load LPH Data
				model.ViewType = "index_restore";

				string lph = _ppLphAppService.GetById(result.ID);
				PPLPHModel lphModel = lph.DeserializeToPPLPH();
				model.LPH = lphModel;

				string extra = _ppLphExtrasAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
				List<PPLPHExtrasModel> extraList = extra.DeserializeToPPLPHExtrasList();
				for (int i = 0; i < extraList.Count; i++) extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
				model.Extras = extraList.OrderBy(x => x.ID).ToList();

				string compo = _ppLphComponentsAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
				IEnumerable<PPLPHComponentsModel> compoList = compo.DeserializeToPPLPHComponentList();
				compoList = compoList.OrderBy(x => x.ID).ToList();

				model.CompoVal = new List<PPLPHCompoValModel>();

				if (compoList != null)
					foreach (var component in compoList.ToList())
					{
						string value = _ppLphValuesAppService.GetBy("LPHComponentID", component.ID, true);
						PPLPHValuesModel valueModel = value.DeserializeToPPLPHValue();

						var CompoVal = new PPLPHCompoValModel();
						CompoVal.Component = component;
						CompoVal.Value = valueModel;
						CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";

						model.CompoVal.Add(CompoVal);
					}
				#endregion
			}

			return View("index", model);
		}

		[HttpPost]
		public ActionResult getPSSData(String BatchID)
		{
			try
			{
				BatchID = BatchID.Trim();
				string strConString = ConfigurationManager.ConnectionStrings["AdoPSS4Conn"].ConnectionString;
				var queryAle = @"SELECT [EventDateTimeLocal],[EventType],[DataName],[Value] FROM [NpssReportBatchData]
                    WHERE EventID = (SELECT [BatchID] FROM [PSS4_Runtime].[dbo].[NpssReportBatch] WHERE BatchIdent = '" + BatchID + @"' and starttimelocal >= starttimelocal - 100) and 
                    ((EventType = 'ST CutClove' and DataName in ('L400 destinationSilo','Batch ID')) or 
                    (EventType = 'END CutClove' and DataName in ('Batch ID','Avg CutWidth','Sampling1_CutWidthCP','Sampling2_CutWidthCP','Sampling3_CutWidthCP','Sampling4_CutWidthCP','Sampling5_CutWidthCP')) or
                    (EventType = 'END CutClove' and DataName in ('Sampling6_CutWidthCP','Sampling7_CutWidthCP','Sampling8_CutWidthCP','Sampling9_CutWidthCP','Sampling10_CutWidthCP')) or
                    (EventType = 'END DryCloveP' and DataName in ('Dryer1TobFlowAverage','Dryer2TobFlowAverage','Dryer1BatchTobTempAvg','Dryer2BatchTobTempAvg','Dryer1InletOVAverage','Dryer2InletOVAverage','Final-OV-Average')) or 
                    (EventType = 'END DryClove' and DataName in ('WeigherTotalizer Z040003','I04080 avgMoisture','WeigherTotalizer Z040007')) or 
                    (EventType = 'End Clove_Link' and DataName in ('WeigherTotalizer Z045001','WeigherTotalizer Z045002'))) ";

				if (AccountProdCenterID == 4 || AccountProdCenterID == 5)
				{
					strConString = ConfigurationManager.ConnectionStrings["AdoPSS4Conn2"].ConnectionString;
					queryAle = @"SELECT [EventDateTimeLocal],[EventType],[DataName],[Value] FROM [NpssReportBatchData]
                                WHERE EventID = ( SELECT [BatchID] FROM [NpssReportBatch] WHERE BatchIdent ='" + BatchID + @"' and starttimelocal >= starttimelocal - 100)and (
                                    (EventType = 'ST Inf_CL' and DataName in ('Blend Code','Batch ID')) 
	                                or (EventType = 'ST CL_CutDry' and DataName in ('CloveCutDry_46135_flowSP','CloveCutDry_46140_flowSP','CloveCutDry_destinationSilo','CloveCutDry_avgTempOut','CloveCutDry_avgAirTemp'))
	                                or (EventType = 'END CL_CutDry' and DataName in ('CloveCutDry_cuttingWidth','CloveCutDry_46030_metalDetCntr','CloveCutDry_46045_metalDetCntr','V3_Cutter-Breakdown-Count','CloveCutDry_46150avgAirTemp','CloveCutDry_46151avgAirTemp','CloveCutDry_1_TobTempAvgTot','CloveCutDry_2_TobTempAvgTot','CloveCutDry_46245_weigherTotalizer','CloveCutDry_destinationSilo','CloveCutDry_TotalWeightInput','CloveCutDry_46025_flowWith1CutSP','CloveCutDry_FinalOVAvg','2ndBatchID','CloveCutDry_1_HoodPressAvg','CloveCutDry_2_HoodPressAvg','CloveCutDry_1_WetTobQty','CloveCutDry_46150_inMoist_16','CloveCutDry_1_ExitOVAvg',	'CloveCutDry_2_WetTobQty','CloveCutDry_46151_inMoist_21','CloveCutDry_2_ExitOVAvg',	'Batch Size'))
	                                or (EventType = 'END CL_PrebSilo' and DataName in ('ClovePreblendSilo_destinationSilo'))
                                )ORDER BY EventDateTimeLocal DESC";
				}

				DataSet dset = new DataSet();
				using (SqlConnection con = new SqlConnection(strConString))
				{

					SqlCommand cmd = new SqlCommand(queryAle, con);
					SqlDataAdapter da = new SqlDataAdapter(cmd);
					da.Fill(dset);

					Helper.LogErrorMessage("dset count :" + dset.Tables[0].Rows.Count, Server.MapPath("~/Uploads/"));
				}
				var data = DataTableToJSONWithJavaScriptSerializer(dset.Tables[0]);

				return Json(new { Status = "True", Data = data }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
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

		// GET: LPHPrimaryCloveCutDryPackingController/Details/5
		public ActionResult Details(int id)
		{
			return View();
		}

		// GET: LPHPrimaryCloveCutDryPackingController/Create
		public ActionResult Create()
		{
			return View();
		}

		// POST: LPHPrimaryCloveCutDryPackingController/Create
		[HttpPost]
		public ActionResult Create(List<PPLPHExtrasModel> detailsExtras, List<PPLPHValuesModel> detailValue, List<PPLPHComponentsModel> detailComponent)
		{
			try
			{
				var lphID = CreateLPHPrimary(
					_ppLphAppService,
					_ppLphLocationsAppService,
					_ppLphExtrasAppService,
					_ppLphSubmissionsAppService,
					_ppLphApprovalAppService,
					_ppLphComponentsAppService,
					_ppLphValuesAppService,
					detailsExtras, detailComponent, detailValue, GetType().Name);

				return Json(new { Status = "True", LPHID = lphID }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		// GET: LPHPrimaryCloveCutDryPackingController/Edit/5
		public ActionResult Edit(long lphid)
		{
			GetReferenceTarget();

			#region Load Header

			LoadLPHHeaderPP(_locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _jobTitleAppService);

			#endregion

			#region Load LPH Data
			var model = new PPLPHEditModel();
			model.ViewType = "edit";

			string lph = _ppLphAppService.GetById(lphid, true);
			PPLPHModel lphModel = lph.DeserializeToPPLPH();
			model.LPH = lphModel;

			string submit = _ppLphSubmissionsAppService.GetBy("LPHID", lphModel.ID);
			LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

			string approval = _ppLphApprovalAppService.GetAll(true);
			List<PPLPHApprovalsModel> approvalList = approval.DeserializeToPPLPHApprovalList().OrderBy(x => x.ID).ToList();
			approvalList = approvalList.Where(x => x.LPHSubmissionID == submitModel.ID && x.UserEmployeeID == AccountEmployeeID).OrderByDescending(x => x.ID).ToList();
			PPLPHApprovalsModel apprModel = approvalList.FirstOrDefault();

			if (apprModel != null && apprModel.Status != null && (apprModel.Status.Trim().ToLower() == "submitted" || apprModel.Status.Trim().ToLower() == "approved"))
				return RedirectToAction("Submitted", new { lphid = lphModel.ID });
			if (submitModel.UserID != AccountID)
				return RedirectToAction("Submitted", new { lphid = lphModel.ID });

			string extra = _ppLphExtrasAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			List<PPLPHExtrasModel> extraList = extra.DeserializeToPPLPHExtrasList();
			for (int i = 0; i < extraList.Count; i++) extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
			model.Extras = extraList.OrderBy(x => x.ID).ToList();

			string compo = _ppLphComponentsAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			IEnumerable<PPLPHComponentsModel> compoList = compo.DeserializeToPPLPHComponentList();
			compoList = compoList.OrderBy(x => x.ID).ToList();

			model.CompoVal = new List<PPLPHCompoValModel>();
			if (compoList != null)
				foreach (var component in compoList.ToList())
				{
					string value = _ppLphValuesAppService.GetBy("LPHComponentID", component.ID, true);
					PPLPHValuesModel valueModel = value.DeserializeToPPLPHValue();

					var CompoVal = new PPLPHCompoValModel();
					CompoVal.Component = component;
					CompoVal.Value = valueModel;
					CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";

					model.CompoVal.Add(CompoVal);
				}
			#endregion

			return View("index", model);
		}

		// POST: LPHPrimaryCloveCutDryPackingController/Edit/5
		[HttpPost]
		public async Task<ActionResult> Edit(long id, List<PPLPHExtrasModel> detailsExtras, List<PPLPHComponentsModel> detailComponent, List<PPLPHValuesModel> detailValue, int isSubmit = 0)
		{
			try
			{
				await EditLPHPrimary(_userAppService, _employeeAppService, _ppLphSubmissionsAppService, _ppLphExtrasAppService, _ppLphComponentsAppService, _ppLphValuesAppService, _ppLphValueHistoriesAppService, _ppLphApprovalAppService, id, detailsExtras, detailValue, isSubmit);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		// GET: LPHPrimaryCloveCutDryPackingController/Delete/5
		public ActionResult Delete(int id)
		{
			return View();
		}

		// POST: LPHPrimaryCloveCutDryPackingController/Delete/5
		[HttpPost]
		public ActionResult Delete(int id, FormCollection collection)
		{
			try
			{
				// TODO: Add delete logic here

				return RedirectToAction("Index");
			}
			catch
			{
				return View();
			}
		}

		#region Approval Part

		public ActionResult Approval(long lphid)
		{
			string submit = _ppLphSubmissionsAppService.GetBy("LPHID", lphid);
			LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

			string approval = _ppLphApprovalAppService.GetAll(true);
			List<PPLPHApprovalsModel> approvalList = approval.DeserializeToPPLPHApprovalList().OrderBy(x => x.ID).ToList();
			approvalList = approvalList.Where(x => x.LPHSubmissionID == submitModel.ID && x.UserEmployeeID == AccountEmployeeID).OrderByDescending(x => x.ID).ToList();
			PPLPHApprovalsModel apprModel = approvalList.FirstOrDefault();

			if (apprModel != null && apprModel.Status.Trim().ToLower() == "approved")
				return RedirectToAction("Submitted", new { lphid = lphid });

			GetReferenceTarget();

			#region Load LPH Data
			var model = new PPLPHEditModel();
			model.ViewType = "approval";

			string lph = _ppLphAppService.GetById(lphid, true);
			PPLPHModel lphModel = lph.DeserializeToPPLPH();
			model.LPH = lphModel;

			string extra = _ppLphExtrasAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			List<PPLPHExtrasModel> extraList = extra.DeserializeToPPLPHExtrasList();
			for (int i = 0; i < extraList.Count; i++) extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
			model.Extras = extraList.OrderBy(x => x.ID).ToList();

			string compo = _ppLphComponentsAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			IEnumerable<PPLPHComponentsModel> compoList = compo.DeserializeToPPLPHComponentList();
			compoList = compoList.OrderBy(x => x.ID).ToList();

			model.CompoVal = new List<PPLPHCompoValModel>();
			if (compoList != null)
				foreach (var component in compoList.ToList())
				{
					string value = _ppLphValuesAppService.GetBy("LPHComponentID", component.ID, true);
					PPLPHValuesModel valueModel = value.DeserializeToPPLPHValue();

					var CompoVal = new PPLPHCompoValModel();
					CompoVal.Component = component;
					CompoVal.Value = valueModel;
					CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";

					model.CompoVal.Add(CompoVal);
				}
			#endregion

			#region Load Header

			LoadLPHHeaderPP(_locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _jobTitleAppService, model.LPH.LocationID);

			#endregion

			return View("index", model);
		}

		[HttpPost]
		public async Task<ActionResult> Approval(long id, string approve, string comment, List<PPLPHExtrasModel> detailsExtras, List<PPLPHComponentsModel> detailComponent, List<PPLPHValuesModel> detailValue)
		{
			try
			{
				await ApproveLPHPrimary(
					_userAppService,
					_employeeAppService,
					_ppLphSubmissionsAppService,
					_ppLphExtrasAppService,
					_ppLphComponentsAppService,
					_ppLphValuesAppService,
					_ppLphValueHistoriesAppService,
					_ppLphApprovalAppService,
					id, detailsExtras, detailValue, comment);

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

				await RejectPPLPH(_userAppService, _employeeAppService, _ppLphSubmissionsAppService, _ppLphApprovalAppService, id, comment);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}
		#endregion

		#region Prev & Submitted

		[HttpPost]
		public ActionResult CheckPrev(long lphid = 0)
		{
			try
			{
				if (lphid != 0)
				{
					string submissions = _ppLphSubmissionsAppService.FindBy("LPHHeader", "LPHPrimaryCloveCutDryPackingController", true); //ambil yang sudah submit saja
					if (!string.IsNullOrEmpty(submissions))
					{
						List<PPLPHSubmissionsModel> submisList = submissions.DeserializeToPPLPHSubmissionsList();
						submisList = submisList.Where(x => x.LPHID < lphid && x.LocationID == AccountLocationID).OrderByDescending(x => x.ID).ToList();

						if (submisList.Count() > 0)
							return Json(new { Status = true, LPHID = submisList.First().LPHID }, JsonRequestBehavior.AllowGet);
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
			GetReferenceTarget();

			#region Load LPH Data
			var model = new PPLPHEditModel();
			model.ViewType = "submitted";

			if (type == 1)
				model.ViewType = "submitted_prev";

			string lph = _ppLphAppService.GetById(lphid, true);
			PPLPHModel lphModel = lph.DeserializeToPPLPH();
			model.LPH = lphModel;

			string extra = _ppLphExtrasAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			List<PPLPHExtrasModel> extraList = extra.DeserializeToPPLPHExtrasList();
			for (int i = 0; i < extraList.Count; i++) extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
			model.Extras = extraList.OrderBy(x => x.ID).ToList();

			string compo = _ppLphComponentsAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			IEnumerable<PPLPHComponentsModel> compoList = compo.DeserializeToPPLPHComponentList();
			compoList = compoList.OrderBy(x => x.ID).ToList();

			model.CompoVal = new List<PPLPHCompoValModel>();
			if (compoList != null)
				foreach (var component in compoList.ToList())
				{
					string value = _ppLphValuesAppService.GetBy("LPHComponentID", component.ID, true);
					PPLPHValuesModel valueModel = value.DeserializeToPPLPHValue();

					var CompoVal = new PPLPHCompoValModel();
					CompoVal.Component = component;
					CompoVal.Value = valueModel;
					CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";

					model.CompoVal.Add(CompoVal);
				}
			#endregion

			#region Load Header
			LoadLPHHeaderPP(_locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _jobTitleAppService, model.LPH.LocationID);
			#endregion

			return View("index", model);
		}

		#endregion

		#region Helper Get Info

		private PPLPHModel GetLPH(long id)
		{
			string data = _ppLphAppService.GetById(id, true);
			PPLPHModel ppLphModel = data.DeserializeToPPLPH();

			return ppLphModel;
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
				string lphSubmission = _ppLphSubmissionsAppService.GetBy("LPHID", lphid, true);
				PPLPHSubmissionsModel submissionsModel = lphSubmission.DeserializeToPPLPHSubmissions();

				string approval = _ppLphApprovalAppService.GetAll(true);
				List<PPLPHApprovalsModel> approvalList = approval.DeserializeToPPLPHApprovalList().OrderBy(x => x.Date).ToList();

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

		private string GetFullName(long userID)
		{
			string user = _userAppService.GetById(userID);
			UserModel userModel = user.DeserializeToUser();

			string emp = _employeeAppService.GetBy("EmployeeID", userModel.EmployeeID, true);
			EmployeeModel empModel = emp.DeserializeToEmployee();

			return empModel.FullName;
		}
		#endregion

	}
}
