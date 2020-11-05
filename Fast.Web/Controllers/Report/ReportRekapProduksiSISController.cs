using Fast.Application.Interfaces;
using Fast.Web.Models.LPH;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models.Report;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
	public class ReportRekapProduksiSISController : BaseController<LPHModel>
	{
		private readonly ILPHAppService _lphAppService;
		private readonly ILocationAppService _locationAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly IWppAppService _wppAppService;
		private readonly ILPHComponentsAppService _lphComponentsAppService;
		private readonly ILPHValuesAppService _lphValuesAppService;
		private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
		private readonly ILPHApprovalsAppService _lphApprovalAppService;
		private readonly IUserAppService _userAppService;

		private readonly ILoggerAppService _logger;

		public ReportRekapProduksiSISController(
			ILPHAppService lphAppService,
			IReferenceAppService referenceAppService,
			ILocationAppService locationAppService,
			IWppAppService wppAppService,
			ILPHComponentsAppService lphComponentsAppService,
			ILPHValuesAppService lphValuesAppService,
			ILPHSubmissionsAppService lPHSubmissionsAppService,
			ILPHApprovalsAppService lphApprovalsAppService,
			IUserAppService userAppService,

			ILoggerAppService logger)
		{
			_lphAppService = lphAppService;
			_referenceAppService = referenceAppService;
			_locationAppService = locationAppService;
			_wppAppService = wppAppService;
			_lphComponentsAppService = lphComponentsAppService;
			_lphValuesAppService = lphValuesAppService;
			_lphSubmissionsAppService = lPHSubmissionsAppService;
			_lphApprovalAppService = lphApprovalsAppService;
			_userAppService = userAppService;

			_logger = logger;
		}

		public ActionResult Index()
		{
			ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

			return View();
		}

		[HttpPost]
		public ActionResult GetReportByParam(string startDate, string endDate, string week, long locationID)
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

				if (startDate != "" && endDate != "")
				{
					DateTime startDt = DateTime.Parse(startDate);
					DateTime endDt = DateTime.Parse(endDate);
				}

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}

		}

		public ActionResult RPH1(string StartDate = "", string EndDate = "", long location = 0, string machine = "")
		{
			var model = new RekapProduksiRPH1Model();

			DateTime baseDate = DateTime.Today;

			var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
			var sunday = monday.AddDays(6);

			//untuk sementara pakai ini karena data terbatas
			monday = DateTime.Parse("2020-02-03");
			sunday = DateTime.Parse("2020-02-09");

			model.StartDate = monday;
			model.EndDate = sunday;

			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("Date", model.StartDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
			filters.Add(new QueryFilter("Date", model.EndDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
			filters.Add(new QueryFilter("IsDeleted", "0"));

			string wpp = _wppAppService.Find(filters);
			model.wppList = wpp.DeserializeToWppList();
			model.wppSimpleList = wpp.DeserializeToWppSimpleList().Distinct().ToList();
			model.DataLPH = new List<RekapProduksiDataLPHModel>();


			//pake filter yg sama; ternyata filter direset setelah eksekusi
			filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("Date", model.StartDate.ToString("yyyy-MM-dd"), Operator.GreaterThanOrEqual));
			filters.Add(new QueryFilter("Date", model.EndDate.ToString("yyyy-MM-dd"), Operator.LessThanOrEqualTo));
			filters.Add(new QueryFilter("IsDeleted", "0"));

			string submissionList = _lphSubmissionsAppService.Find(filters);
			List<LPHSubmissionsModel> submissions = submissionList.DeserializeToLPHSubmissionsList();

			//hanya ambil yang sudah approved
			submissions = submissions.Where(x => x.IsComplete == true).ToList();

			//ambil packer saja?
			submissions = submissions.Where(x => x.LPHHeader == "PackerController").ToList();

			//untuk sementara abaikan dulu
			List<long> locations = new List<long>();

			var LocatSubs = new List<ProdCenterSubsModel>();
			var ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
			foreach (var locat in ProductionCenterList)
			{
				if (locat.Value != "0" && (location == 0 || (location.ToString() == locat.Value)))
				{
					var ProdCenterSubs = new ProdCenterSubsModel();

					ProdCenterSubs.LocationID = Int64.Parse(locat.Value);
					ProdCenterSubs.Subs = _locationAppService.GetLocIDListByLocType(ProdCenterSubs.LocationID, "productioncenter");
					locations.AddRange(ProdCenterSubs.Subs);

					LocatSubs.Add(ProdCenterSubs);
				}
			}
			submissions = submissions.Where(x => locations.Contains(x.LocationID)).ToList();



			if (submissions.Count() > 0)
			{
				/*
                //ambil by min & max ID; sepertinya lebih cepat drpd find berkali2, benar tak?
                var minSubmissionID = submissions.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
                var maxSubmissionID = submissions.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

                filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("IsDeleted", "0"));
                filters.Add(new QueryFilter("LPHSubmissionID", minSubmissionID.ToString(), Operator.GreaterThanOrEqual));
                filters.Add(new QueryFilter("LPHSubmissionID", maxSubmissionID.ToString(), Operator.LessThanOrEqualTo));
                filters.Add(new QueryFilter("Status", "          ", Operator.NotEqual)); //barangkali bisa lumayan meringankan, semoga sama semua pakai nvarchar
                filters.Add(new QueryFilter("Status", "Submitted ", Operator.NotEqual));
                filters.Add(new QueryFilter("Status", "Draft     ", Operator.NotEqual)); 

                string approvals = _lphApprovalAppService.Find(filters);
                List<LPHApprovalsModel> approvalList = approvals.DeserializeToLPHApprovalList();

                */

				//setelah difilter, cari pasangan LPH-nya; sebisa mungkin hindari getall
				foreach (var item in submissions.ToList())
				{
					string lphM = _lphAppService.GetById(item.LPHID); //harusnya getbyid lebih cepat drpd find where lphid >= lowestID
					LPHModel lphModel = lphM.DeserializeToLPH();

					// chanif: exclude LPH yang sudah dihapus
					if (lphModel == null)
					{
						submissions.Remove(item);
						continue;
					}
					else if (lphModel.IsDeleted)
					{
						submissions.Remove(item);
						continue;
					}

					/*
                    LPHApprovalsModel approvalModel = approvalList.Where(x => x.LPHSubmissionID == item.ID && x.Status.Trim().ToLower() == "approved").LastOrDefault();
                    if (approvalModel != null && approvalModel.ID > 0)
                    {
                        // hanya hitung data yg sudah approved
                    }
                    else
                    {
                        // aneh, buang aja drpd error
                        submissions.Remove(item);
                        continue;
                    }
                    */

					item.LPHHeader = item.LPHHeader.Replace("Controller", "");
				}

				if (submissions.Count() > 0)
				{
					var LPHIDs = submissions.Select(x => x.LPHID).ToList();
					var minLPHID = submissions.DefaultIfEmpty().Min(x => x == null ? 0 : x.LPHID);
					var maxLPHID = submissions.DefaultIfEmpty().Max(x => x == null ? 0 : x.LPHID);

					filters = new List<QueryFilter>();
					filters.Add(new QueryFilter("IsDeleted", "0"));
					filters.Add(new QueryFilter("LPHID", minLPHID.ToString(), Operator.GreaterThanOrEqual));
					filters.Add(new QueryFilter("LPHID", maxLPHID.ToString(), Operator.LessThanOrEqualTo));

					string comps = _lphComponentsAppService.Find(filters);
					List<LPHComponentsModel> componentList = comps.DeserializeToLPHComponentList();

					// exclude data yg ndak masuk di submission
					componentList = componentList.Where(x => LPHIDs.Contains(x.LPHID)).ToList();

					if (componentList.Count() > 0)
					{
						// populate values
						var compoIDs = componentList.Select(x => x.ID).ToList();
						var minCompoID = componentList.DefaultIfEmpty().Min(x => x == null ? 0 : x.ID);
						var maxCompoID = componentList.DefaultIfEmpty().Max(x => x == null ? 0 : x.ID);

						filters = new List<QueryFilter>();
						filters.Add(new QueryFilter("IsDeleted", "0"));
						filters.Add(new QueryFilter("LPHComponentID", minCompoID.ToString(), Operator.GreaterThanOrEqual));
						filters.Add(new QueryFilter("LPHComponentID", maxCompoID.ToString(), Operator.LessThanOrEqualTo));

						string values = _lphValuesAppService.Find(filters);
						List<LPHValuesModel> valueList = values.DeserializeToLPHValueList();

						// exclude value yg ndak masuk di componentList
						valueList = valueList.Where(x => compoIDs.Contains(x.LPHComponentID)).ToList();

						// header untuk component
						var compoNames = componentList.Select(x => x.ComponentName).Distinct().ToList();
						foreach (var item in compoNames.ToList())
						{
							var content = item.Replace("generalInfo-", "").Replace("TeamLeader", "Supervisor");
							if (String.IsNullOrWhiteSpace(content))
							{
								compoNames.Remove(item);
								continue;
							}
							//Sheet.Cells[colHeader, rowHeader++].Value = content;
						}

						foreach (var item in submissions)
						{
							var dataLPH = new RekapProduksiDataLPHModel();

							foreach (var name in compoNames)
							{
								dynamic content = "";

								var compo = componentList.Where(x => x.LPHID == item.LPHID && x.ComponentName == name).FirstOrDefault();
								if (compo != null && compo.ID != 0)
								{
									var value = valueList.Where(x => x.LPHComponentID == compo.ID).FirstOrDefault();

									if (value != null && !String.IsNullOrEmpty(value.Value))
									{
										if (value.ValueType.Trim() == "Numeric")
										{
											if (value.Value.Contains("."))
											{
												//tak anggap double
												double number = 0;
												Double.TryParse(value.Value, out number);
												content = number;
											}
											else
											{
												//tak anggep long
												Int64 number = 0;
												Int64.TryParse(value.Value, out number);
												content = number;
											}
										}
										else if (value.ValueType.Trim() == "ImageURL")
										{
											if (value.Value == "_no_image.png")
											{
												content = "no image";
											}
											else
											{
												content = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"] + "/Uploads/lph/" + item.LPHHeader.ToLower() + "/" + value.Value;
											}
										}
										else
										{
											content = value.Value;
										}
									}
								}

								//Sheet.Cells[colContent, rowContent++].Value = content;
							}

							model.DataLPH.Add(dataLPH);
						}


						SetTrueTempData(UIResources.ExtractSuccess);
					}
					else
					{
						SetFalseTempData(UIResources.NoDataInSelectedCriteria);
					}
				}
				else
				{
					SetFalseTempData(UIResources.NoDataInSelectedCriteria);
				}
			}



			return PartialView(model);
		}

		[HttpPost]
		public ActionResult GenerateExcel(DateTime dtFilter, long prodCenterID)
		{
			try
			{
				//byte[] excelData = ExcelGenerator.ExportRekapProduksi(AccountName);

				//Response.Clear();
				//Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				//Response.AddHeader("content-disposition", "attachment;filename=RekapProduksi.xlsx");
				//Response.BinaryWrite(excelData);
				//Response.End();
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.GenerateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}
		#region Helper

		[HttpPost]
		public int GetCurrentWeekNumber(string date)
		{
			DateTime dt = DateTime.Parse(date);
			var weeknum = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
			return weeknum;
		}

		[HttpPost]
		public ActionResult GetDepartmentByProdCenterID(long id)
		{
			List<SelectListItem> _menuList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, id);

			return Json(_menuList, JsonRequestBehavior.AllowGet);
		}
		#endregion

	}
}
