using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
	[CustomAuthorize("master")]
	public class ReferenceController : BaseController<ReferenceModel>
	{
		private readonly IReferenceAppService _referenceAppService;
		private readonly IReferenceDetailAppService _referenceDetailAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;
		private readonly ICalendarHolidayAppService _calHolidayAppService;

		public ReferenceController(
			IReferenceAppService referenceAppService,
			ILoggerAppService logger,
			IMenuAppService menuService,
			ICalendarHolidayAppService calHolidayAppService,
			IReferenceDetailAppService referenceDetailAppService)
		{
			_referenceAppService = referenceAppService;
			_referenceDetailAppService = referenceDetailAppService;
			_menuService = menuService;
			_logger = logger;
			_calHolidayAppService = calHolidayAppService;
		}

		// GET: Reference
		public ActionResult Index()
		{
			ViewBag.ReferenceList = DropDownHelper.BindDropDownReference(_referenceAppService);

			ReferenceTreeModel model = GetTreeReference();
			model.Access = GetAccess(WebConstants.MenuSlug.REFERENCE, _menuService);

			return View(model);
		}
		private ReferenceTreeModel GetTreeReference()
		{
			ReferenceTreeModel model = new ReferenceTreeModel();

			// Getting all data    			
			string referenceList = _referenceAppService.GetAll(true);
			List<ReferenceModel> references = referenceList.DeserializeToReferenceList();

			// exclude / hide location master in reference
			references = references.Where(
				x => !x.Name.Equals("PC") &&
				!x.Name.Equals("Dep") &&
				!x.Name.Equals("SubDep") &&
                !x.Name.Equals("Brand") &&
                !x.Name.Equals("Blend") &&
                !x.Name.Equals("LT")).OrderBy(x => x.Name).ToList();

			// Construct reference details			
			foreach (var item in references)
			{
				item.ReferenceDetails = GetReferenceDetails(item.ID);
			}

			// add parent list
			model.Parents.AddRange(references);

			return model;
		}

		[HttpPost]
		public ActionResult Add(ReferenceTreeModel model)
		{
			try
			{
				ViewBag.ReferenceList = DropDownHelper.BindDropDownReference(_referenceAppService);
				model.Access = GetAccess(WebConstants.MenuSlug.REFERENCE, _menuService);

				if (!ModelState.IsValid)
				{
					return View(model);
				}

				// map new location
				if (model.ReferenceID == 0)
					AddNewReference(model);
				else
					AddNewReferenceDetail(model);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		private void AddNewReferenceDetail(ReferenceTreeModel inputModel)
		{
			ReferenceDetailModel model = new ReferenceDetailModel();
			model.ReferenceID = inputModel.ReferenceID;
			model.Code = inputModel.Code;
			model.Description = inputModel.Description;
			model.ModifiedBy = AccountName;
			model.ModifiedDate = DateTime.Now;

			string data = JsonHelper<ReferenceDetailModel>.Serialize(model);

			_referenceAppService.AddDetail(data);
		}

		private void AddNewReference(ReferenceTreeModel inputModel)
		{
			ReferenceModel model = new ReferenceModel();
			model.Name = inputModel.Name;
			model.Purpose = inputModel.Purpose;
			model.ModifiedBy = AccountName;
			model.ModifiedDate = DateTime.Now;

			string data = JsonHelper<ReferenceModel>.Serialize(model);

			_referenceAppService.Add(data);
		}

		// POST: Reference/Create
		[HttpPost]
		public ActionResult Create(string name, string purpose, string typeId, List<ReferenceDetailModel> details)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
				}

				if (typeId == "0" && (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(purpose)))
				{
					return Json(new { Status = "False", ErrorMessage = UIResources.NamePurposeAreMandatory }, JsonRequestBehavior.AllowGet);
				}

				if (typeId != "0")
				{
					string reference = _referenceAppService.GetById(long.Parse(typeId), true);
					ReferenceModel refModel = reference.DeserializeToReference();
					if ((refModel.Name != name && !string.IsNullOrEmpty(name)) ||
						(refModel.Purpose != purpose && !string.IsNullOrEmpty(purpose)))
					{
						refModel.Name = name;
						refModel.Purpose = purpose;
						refModel.ModifiedBy = AccountName;
						refModel.ModifiedDate = DateTime.Now;

						string updatedRef = JsonHelper<ReferenceModel>.Serialize(refModel);
						_referenceAppService.Update(updatedRef);
					}

					string refDetails = _referenceDetailAppService.FindByNoTracking("ReferenceID", typeId, true);
					List<ReferenceDetailModel> detailModelList = refDetails.DeserializeToRefDetailList();

					// delete if any removed detail
					foreach (var item in detailModelList)
					{
						if (!details.Any(x => x.ID == item.ID))
						{
							_referenceDetailAppService.Remove(item.ID);

							if (long.Parse(typeId) == (long)ReferenceEnum.HolidayType)
							{
								string holidays = _calHolidayAppService.FindByNoTracking("HolidayTypeID", item.ID.ToString(), true);
								List<CalendarHolidayModel> holidayList = holidays.DeserializeToCalendarHolidayList();

								foreach (var holiday in holidayList)
								{
									_calHolidayAppService.Remove(holiday.ID);
								}
							}
						}
					}

					if (details != null)
					{
						foreach (var item in details)
						{
							item.ReferenceID = long.Parse(typeId);
							item.ModifiedBy = AccountName;
							item.ModifiedDate = DateTime.Now;

							ReferenceDetailModel exist = detailModelList.Where(x => x.ID == item.ID).FirstOrDefault();							
							if (exist == null)
							{
								string detail = JsonHelper<ReferenceDetailModel>.Serialize(item);
								_referenceDetailAppService.Add(detail);
							}
							else
							{
								string detail = JsonHelper<ReferenceDetailModel>.Serialize(item);
								_referenceDetailAppService.Update(detail);

								if (long.Parse(typeId) == (long)ReferenceEnum.HolidayType && exist.Code != item.Code)
								{
									string holidays = _calHolidayAppService.FindByNoTracking("HolidayTypeID", item.ID.ToString(), true);
									List<CalendarHolidayModel> holidayList = holidays.DeserializeToCalendarHolidayList();

									foreach (var holiday in holidayList)
									{
										holiday.Color = item.Code;
										holiday.ModifiedBy = AccountName;
										holiday.ModifiedDate = DateTime.Now;

										string hol = JsonHelper<CalendarHolidayModel>.Serialize(holiday);
										_calHolidayAppService.Update(hol);
									}
								}
							}
						}
					}
				}
				else
				{
					ReferenceModel model = new ReferenceModel { Name = name, Purpose = purpose };

					string exist = _referenceAppService.GetBy("Name", model.Name);
					if (!string.IsNullOrEmpty(exist))
					{
						return Json(new { Status = "False", ErrorMessage = string.Format(UIResources.DataExist, "Reference", model.Name) }, JsonRequestBehavior.AllowGet);
					}

					if (details == null)
					{
						return Json(new { Status = "False", ErrorMessage = string.Format(UIResources.ReferenceMissingDetails, "Reference", model.Name) }, JsonRequestBehavior.AllowGet);
					}

					model.ModifiedBy = AccountName;
					model.ModifiedDate = DateTime.Now;

					string data = JsonHelper<ReferenceModel>.Serialize(model);

					_referenceAppService.Add(data);

					string newreference = _referenceAppService.GetBy("Name", model.Name);

					ReferenceModel newref = newreference.DeserializeToReference();

					foreach (var item in details)
					{
						item.ReferenceID = newref.ID;
						item.ModifiedBy = AccountName;
						item.ModifiedDate = DateTime.Now;
						string detail = JsonHelper<ReferenceDetailModel>.Serialize(item);
						_referenceDetailAppService.Add(detail);
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

		public ActionResult GetReferenceDetail(string referenceID)
		{
			string referenceDetails = _referenceAppService.FindDetailBy("ReferenceID", referenceID, true);
			List<ReferenceDetailModel> modelList = referenceDetails.DeserializeToRefDetailList();

			return Json(modelList, JsonRequestBehavior.AllowGet);
		}

		// GET: Reference/Edit/5
		public ActionResult Edit(int id)
		{
			string refDetailObject = _referenceDetailAppService.GetBy("ID", id.ToString(), true);
			ReferenceDetailModel referenceDetailModel = refDetailObject.DeserializeToRefDetail();

			string reference = _referenceAppService.GetBy("ID", referenceDetailModel.ReferenceID.ToString(), true);
			ReferenceModel referenceModel = reference.DeserializeToReference();

			string refDetailList = _referenceDetailAppService.FindBy("ReferenceID", referenceModel.ID.ToString(), true);
			List<ReferenceDetailModel> referenceDetailsModel = refDetailList.DeserializeToRefDetailList();

			referenceModel.ReferenceDetails = referenceDetailsModel;

			return View(referenceModel);
		}

		// POST: Reference/Edit/5
		[HttpPost]
		public ActionResult Edit(string refID, string name, string purpose, List<ReferenceDetailModel> details)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
				}

				string reference = _referenceAppService.GetById(long.Parse(refID), true);
				ReferenceModel referenceModel = reference.DeserializeToReference();
				if (!referenceModel.Name.Equals(name) || !referenceModel.Purpose.Equals(purpose))
				{
					referenceModel.Name = name;
					referenceModel.Purpose = purpose;
					referenceModel.ModifiedBy = AccountName;
					referenceModel.ModifiedDate = DateTime.Now;
					string updatedReference = JsonHelper<ReferenceModel>.Serialize(referenceModel);
					_referenceAppService.Update(updatedReference);
				}

				// delete the existing details
				string oldDetails = _referenceDetailAppService.FindByNoTracking("ReferenceID", refID, true);
				List<ReferenceDetailModel> referenceDetailsModelList = oldDetails.DeserializeToRefDetailList();

				foreach (var item in referenceDetailsModelList)
				{
					if (!details.Any(x => x.ID == item.ID))
					{
						_referenceDetailAppService.Remove(item.ID);
					}
				}

				// add new details
				foreach (var item in details)
				{
					ReferenceDetailModel temp = referenceDetailsModelList.Where(x => x.ID == item.ID).FirstOrDefault();
					if (temp == null)
					{
						item.ReferenceID = referenceModel.ID;
						item.ModifiedBy = AccountName;
						item.ModifiedDate = DateTime.Now;

						string detail = JsonHelper<ReferenceDetailModel>.Serialize(item);

						_referenceDetailAppService.Add(detail);
					}
					else if (temp.Code != item.Code || temp.Description != item.Description)
					{
						temp.Code = item.Code;
						temp.Description = item.Description;
						temp.ModifiedBy = AccountName;
						temp.ModifiedDate = DateTime.Now;

						string detail = JsonHelper<ReferenceDetailModel>.Serialize(temp);

						_referenceDetailAppService.Update(detail);
					}
				}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

        public ActionResult ExportExcel()
        {
            try
            {
                ReferenceTreeModel model = GetTreeReference();

                byte[] excelData = ExcelGenerator.ExportMasterReference(model, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-Reference.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        // GET: Reference/Delete/5
        public ActionResult Delete(int id)
		{
			return View();
		}

		// POST: Reference/Delete/5
		[HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				ReferenceDetailModel refDetailModel = GetReferenceDetail(id);
				_referenceDetailAppService.Remove(id);

				if (refDetailModel.ReferenceID == (long)ReferenceEnum.HolidayType)
				{
					string holidays = _calHolidayAppService.FindByNoTracking("HolidayTypeID", id.ToString(), true);
					List<CalendarHolidayModel> holidayList = holidays.DeserializeToCalendarHolidayList();

					foreach (var item in holidayList)
					{
						_calHolidayAppService.Remove(item.ID);
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

		[HttpPost]
		public ActionResult DeleteReference(long id)
		{
			try
			{
				ReferenceModel refModel = GetReference(id);
				refModel.IsDeleted = true;
				refModel.ModifiedBy = AccountName;
				refModel.ModifiedDate = DateTime.Now;

				string reference = JsonHelper<ReferenceModel>.Serialize(refModel);
				_referenceAppService.Update(reference);

				//string refDetails = _referenceDetailAppService.FindByNoTracking("ReferenceID", id.ToString(), true);
				//List<ReferenceDetailModel> detailList = refDetails.DeserializeToReferenceDetailList();

				//foreach (var item in detailList)
				//{
				//	item.IsDeleted = true;
				//	item.ModifiedBy = AccountName;
				//	item.ModifiedDate = DateTime.Now;

				//	string detail = JsonHelper<ReferenceDetailModel>.Serialize(item);
				//	_referenceAppService.Update(detail);
				//}

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult GetAll()
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

				// Getting all data    			
				string referenceList = _referenceAppService.GetAll(true);
				List<ReferenceModel> references = referenceList.DeserializeToReferenceList();
				references = references.Where(
					x => !x.Name.Equals("Country") &&
					!x.Name.Equals("PC") &&
					!x.Name.Equals("Dep") &&
					!x.Name.Equals("SubDep") &&
					!x.Name.Equals("LT")).ToList();

				// View Model to make it easier populated in datatables
				List<ReferenceViewModel> viewModels = new List<ReferenceViewModel>();

				// Construct reference details			
				foreach (var item in references)
				{
					item.ReferenceDetails = GetReferenceDetails(item.ID);
					foreach (var detail in item.ReferenceDetails)
					{
						viewModels.Add(new ReferenceViewModel
						{
							ReferenceID = item.ID,
							ReferenceDetailID = detail.ID,
							Name = item.Name,
							Purpose = item.Purpose,
							Code = detail.Code,
							Description = detail.Description
						});
					}
				}

				int recordsTotal = viewModels.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					viewModels = viewModels.Where(m => m.Name.ToLower().Contains(searchValue.ToLower()) ||
												m.Purpose.ToLower().Contains(searchValue.ToLower()) ||
												m.Code.ToLower().Contains(searchValue.ToLower()) ||
												m.Description.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "id":
								viewModels = viewModels.OrderBy(x => x.ReferenceDetailID).ToList();
								break;
							case "name":
								viewModels = viewModels.OrderBy(x => x.Name).ToList();
								break;
							case "purpose":
								viewModels = viewModels.OrderBy(x => x.Purpose).ToList();
								break;
							case "code":
								viewModels = viewModels.OrderBy(x => x.Code).ToList();
								break;
							case "description":
								viewModels = viewModels.OrderBy(x => x.Description).ToList();
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
								viewModels = viewModels.OrderByDescending(x => x.ReferenceDetailID).ToList();
								break;
							case "name":
								viewModels = viewModels.OrderByDescending(x => x.Name).ToList();
								break;
							case "purpose":
								viewModels = viewModels.OrderByDescending(x => x.Purpose).ToList();
								break;
							case "code":
								viewModels = viewModels.OrderByDescending(x => x.Code).ToList();
								break;
							case "description":
								viewModels = viewModels.OrderByDescending(x => x.Description).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = viewModels.Count();

				// Paging 
				var data = viewModels.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<ReferenceViewModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		private List<ReferenceDetailModel> GetReferenceDetails(long referenceID)
		{
			string referenceDetails = _referenceDetailAppService.FindBy("ReferenceID", referenceID.ToString(), true);
			List<ReferenceDetailModel> referenceDetailList = referenceDetails.DeserializeToRefDetailList();
			foreach (var item in referenceDetailList)
			{
				item.ReferenceID = referenceID;
			}

			return referenceDetailList;
		}

		private ReferenceModel GetReference(long referenceID)
		{
			string reference = _referenceAppService.GetById(referenceID, true);
			ReferenceModel referenceModel = reference.DeserializeToReference();

			return referenceModel;
		}

		private ReferenceDetailModel GetReferenceDetail(long refDetailID)
		{
			string refDetail = _referenceDetailAppService.GetById(refDetailID, true);
			ReferenceDetailModel refDetailModel = refDetail.DeserializeToRefDetail();

			return refDetailModel;
		}
	}
}
