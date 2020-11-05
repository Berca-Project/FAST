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
#endregion

namespace Fast.Web.Controllers.LPH
{
	public class MakerTimbanganController : BaseController<LPHModel>
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
		public MakerTimbanganController(
			ILPHAppService lphAppService,
			IWppAppService wppAppService,
			IWeeksAppService weeksAppService,
			ILPHComponentsAppService lphComponentsAppService,
			ILPHLocationsAppService lphLocationsAppService,
			ILPHValuesAppService lphValuesAppService,
			ILPHApprovalsAppService lphApprovalsAppService,
			ILPHValueHistoriesAppService lphValueHistoriesAppService,
			ILPHExtrasAppService lphExtrasAppService,
			ILoggerAppService logger,
			IReferenceAppService referenceAppService,
			IUserAppService userAppService,
			IUserRoleAppService userRoleAppService,
			IMppAppService mppAppService,
			IReferenceDetailAppService referenceDetailAppService,
			IMachineAppService machineAppService,
			IEmployeeAppService employeeAppService,
			ILocationAppService locationAppService,
			IBrandAppService brandAppService,
			ILPHSubmissionsAppService lPHSubmissionsAppService,
			IJobTitleAppService jobTitleAppService)
		{
			_jobTitleAppService = jobTitleAppService;
			_locationAppService = locationAppService;
			_brandAppService = brandAppService;
			_userRoleAppService = userRoleAppService;
			_lphAppService = lphAppService;
			_lphComponentsAppService = lphComponentsAppService;
			_lphLocationsAppService = lphLocationsAppService;
			_lphValuesAppService = lphValuesAppService;
			_lphApprovalAppService = lphApprovalsAppService;
			_lphValueHistoriesAppService = lphValueHistoriesAppService;
			_lphExtrasAppService = lphExtrasAppService;
			_logger = logger;
			_referenceAppService = referenceAppService;
			_userAppService = userAppService;
			_mppAppService = mppAppService;
			_referenceDetailAppService = referenceDetailAppService;
			_weeksAppService = weeksAppService;
			_wppAppService = wppAppService;
			_lphSubmissionsAppService = lPHSubmissionsAppService;
			_machineAppService = machineAppService;
			_employeeAppService = employeeAppService;
		}
		#endregion

		#region ::Public Methods::
		public ActionResult Index()
		{
			LPHHeaderModel model = LPHHeader(_weeksAppService, _mppAppService, _wppAppService);

			LoadLPHHeader(model, _locationAppService, _employeeAppService, _userAppService, _userRoleAppService, _machineAppService, _brandAppService, _jobTitleAppService, _wppAppService);

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

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);

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

			/******* chanif: untuk isi header **************/
			string machine = _machineAppService.GetAll(true);
			List<MachineModel> machineList = machine.DeserializeToMachineList();
			ViewBag.Machines = machineList;

			string brand = _referenceDetailAppService.FindBy("ReferenceID", "3", true);
			List<ReferenceDetailModel> brandList = brand.DeserializeToRefDetailList();
			ViewBag.Brand = brandList;

			string employee = _employeeAppService.GetAll(true);
			List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();

			ViewBag.Prodtechs = employeeList.Where(x => x.PositionDesc == "Production Technician").OrderBy(x => x.FullName).ToList();
			ViewBag.Foremans = employeeList.Where(x => x.PositionDesc == "Foreman").OrderBy(x => x.FullName).ToList();
			ViewBag.Mechanics = employeeList.Where(x => x.PositionDesc == "Mechanic").OrderBy(x => x.FullName).ToList();
			ViewBag.Electrics = employeeList.Where(x => x.PositionDesc == "Electrician").OrderBy(x => x.FullName).ToList();
			ViewBag.Leaders = employeeList.Where(x => x.PositionDesc == "Team Leader").OrderBy(x => x.FullName).ToList();

			/******* chanif: untuk isi header - end **************/

			string lph = _lphAppService.GetById(lphid, true);
			LPHModel lphModel = lph.DeserializeToLPH();
			model.LPH = lphModel;

			string extra = _lphExtrasAppService.FindByNoTracking("LPHID", lphModel.ID.ToString(), true);
			List<LPHExtrasModel> extraList = extra.DeserializeToLPHExtraList();
			for (int i = 0; i < extraList.Count; i++) extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
			model.Extras = extraList.OrderBy(x => x.ID).ToList();

			string compo = _lphComponentsAppService.FindBy("LPHID", lphModel.ID.ToString(), true);
			IEnumerable<LPHComponentsModel> compoList = compo.DeserializeToLPHComponentList();
			compoList = compoList.OrderBy(x => x.ID).ToList();

			model.CompoVal = new List<LPHCompoValModel>();

			string values = _lphValuesAppService.FindBy("SubmissionID", lphid, true);
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

			return View(model);
		}

		[HttpPost]
		public ActionResult Edit(long id, List<LPHExtrasModel> detailExtras, List<LPHComponentsModel> detailComponent, List<LPHValuesModel> detailValue)
		{
			try
			{
				string lph = _lphAppService.GetById(id, true);
				LPHModel lphModel = lph.DeserializeToLPH();

				string submit = _lphSubmissionsAppService.GetBy("LPHID", id.ToString(), true);
				LPHSubmissionsModel submitModel = submit.DeserializeToLPHSubmissions();

				string extra = _lphExtrasAppService.FindByNoTracking("LPHID", id.ToString(), true);
				List<LPHExtrasModel> extraList = extra.DeserializeToLPHExtraList();

				foreach (var ext in extraList)
				{
					if (!ext.IsDeleted)
					{
						ext.IsDeleted = true;

						string extt = JsonHelper<LPHExtrasModel>.Serialize(ext);
						_lphExtrasAppService.Update(extt);
					}
				}

				if (detailExtras != null)
				{
					foreach (var item in detailExtras)
					{
						item.LPHID = id;
						item.Date = DateTime.Now;
						item.UserID = AccountID;
						item.ModifiedBy = AccountName;
						item.ModifiedDate = DateTime.Now;
						item.LocationID = AccountLocationID;
						item.SubShift = 1;

						string detail = JsonHelper<LPHExtrasModel>.Serialize(item);
						_lphExtrasAppService.Add(detail);
					}
				}



				string checkDataComps = _lphComponentsAppService.FindByNoTracking("LPHID", id.ToString(), true);
				List<LPHComponentsModel> dataComps = checkDataComps.DeserializeToLPHComponentList();
				dataComps = dataComps.OrderBy(x => x.ID).ToList();

				if (dataComps.Count() == detailValue.Count())
				{
					for (int i = 0; i < dataComps.Count(); i++)
					{
						var componentID = dataComps[i].ID;

						string value = _lphValuesAppService.FindByNoTracking("LPHComponentID", componentID.ToString(), true);
						List<LPHValuesModel> valueList = value.DeserializeToLPHValueList();
						var valueModel = valueList.FirstOrDefault();


						if (valueModel != null ? valueModel.Value != detailValue[i].Value : detailValue[i].Value != null)
						{
							if (detailValue[i].ValueType == "ImageURL")
							{
								if (detailValue[i].Value != null ? detailValue[i].Value.Length > 200 : false)
								{
									Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
									Random r = new Random();
									var fileName = unixTimestamp.ToString() + "_" + r.Next() + ".jpeg";

									string filePath = Server.MapPath("~/Uploads/lph/makerTimbangan/") + fileName;

									var bytes = Convert.FromBase64String(detailValue[i].Value.Replace("data:image/jpeg;base64,", ""));
									using (var imageFile = new FileStream(filePath, FileMode.Create))
									{
										imageFile.Write(bytes, 0, bytes.Length);
										imageFile.Flush();
									}

									detailValue[i].Value = fileName;
								}
								else
								{
									continue;
								}
							}

							LPHValueHistoriesModel model = new LPHValueHistoriesModel
							{
								LPHValuesID = valueModel.ID,
								OldValue = valueModel.Value,
								NewValue = detailValue[i].Value,
								UserID = AccountID,
								ModifiedBy = AccountName,
								ModifiedDate = DateTime.Now
							};

							string dataHistory = JsonHelper<LPHValueHistoriesModel>.Serialize(model);
							_lphValueHistoriesAppService.Add(dataHistory);

							valueModel.Value = detailValue[i].Value;
							_lphValuesAppService.UpdateModel(valueModel);
						}
					}
				}


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

			/******* chanif: untuk isi header **************/
			string machine = _machineAppService.GetAll(true);
			List<MachineModel> machineList = machine.DeserializeToMachineList();
			ViewBag.Machines = machineList;

			string brand = _referenceDetailAppService.FindBy("ReferenceID", "3", true);
			List<ReferenceDetailModel> brandList = brand.DeserializeToRefDetailList();
			ViewBag.Brand = brandList;

			string employee = _employeeAppService.GetAll(true);
			List<EmployeeModel> employeeList = employee.DeserializeToEmployeeList();

			ViewBag.Prodtechs = employeeList.Where(x => x.PositionDesc == "Production Technician").OrderBy(x => x.FullName).ToList();
			ViewBag.Foremans = employeeList.Where(x => x.PositionDesc == "Foreman").OrderBy(x => x.FullName).ToList();
			ViewBag.Mechanics = employeeList.Where(x => x.PositionDesc == "Mechanic").OrderBy(x => x.FullName).ToList();
			ViewBag.Electrics = employeeList.Where(x => x.PositionDesc == "Electrician").OrderBy(x => x.FullName).ToList();
			ViewBag.Leaders = employeeList.Where(x => x.PositionDesc == "Team Leader").OrderBy(x => x.FullName).ToList();

            /******* chanif: untuk isi header - end **************/

            var model = LPHSPHelper.SetupLPHApprovalModel(_lphAppService, _lphExtrasAppService, _lphComponentsAppService, _lphValuesAppService, submitModel.ID, lphid);

            return View(model);
		}

		[HttpPost]
		public async Task<ActionResult> Approval(long id, string comment, List<LPHExtrasModel> detailExtras, LPHSubmissionsModel detailSubmission, List<LPHComponentsModel> detailComponent, List<LPHValuesModel> detailValue)
		{
			try
			{
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
				string lphSubmission = _lphSubmissionsAppService.GetBy("LPHID", lphid, true);
				LPHSubmissionsModel submissionsModel = lphSubmission.DeserializeToLPHSubmissions();

				string approval = _lphApprovalAppService.GetAll(true);
				List<LPHApprovalsModel> approvalList = approval.DeserializeToLPHApprovalList().OrderBy(x => x.Date).ToList();

				approvalList = approvalList.Where(x => x.LPHSubmissionID == submissionsModel.ID).ToList();
				if (approvalList.Count > 0)
				{
					if (approvalList.Count == 1)
					{
						approvalList.ElementAt(0).User = GetFullName(approvalList.ElementAt(0).UserID);
						approvalList.ElementAt(0).Role = "Requestor";
					}
					if (approvalList.Count == 2)
					{
						approvalList.ElementAt(0).User = GetFullName(approvalList.ElementAt(0).UserID);
						approvalList.ElementAt(0).Role = "Requestor";
						approvalList.ElementAt(1).User = GetFullName(approvalList.ElementAt(1).UserID);
						approvalList.ElementAt(1).Role = "Approver";

					}
					if (approvalList.Count == 3)
					{
						approvalList.ElementAt(0).User = GetFullName(approvalList.ElementAt(0).UserID);
						approvalList.ElementAt(0).Role = "Requestor";
						approvalList.ElementAt(1).User = GetFullName(approvalList.ElementAt(1).UserID);
						approvalList.ElementAt(1).Role = "Approver";
						approvalList.ElementAt(2).User = GetFullName(approvalList.ElementAt(2).UserID);
						approvalList.ElementAt(2).Role = "Approver";
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
		#endregion
	}
}
