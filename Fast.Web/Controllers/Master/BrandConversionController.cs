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
	public class BrandConversionController : BaseController<BrandConversionModel>
	{		
		private readonly IBrandConversionAppService _brandConversionAppService;
		private readonly IBrandAppService _brandAppService;
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;

		public BrandConversionController(
			IBrandConversionAppService BrandConversionService,
			IBrandAppService brandAppService,
			IReferenceAppService referenceAppService,
			IMenuAppService menuService,
			ILoggerAppService logger)
		{			
			_brandConversionAppService = BrandConversionService;
			_brandAppService = brandAppService;
			_referenceAppService = referenceAppService;
			_menuService = menuService;
			_logger = logger;
		}

		// GET: BrandConversion
		public ActionResult Index()
		{
			GetTempData();

			BrandConversionModel model = new BrandConversionModel();
			model.Access = GetAccess(WebConstants.MenuSlug.BRAND_CONVERSION, _menuService);

			return View(model);
		}

		public ActionResult GenerateExcel()
		{
			try
			{
				// Getting all data    			
				string brandConversions = _brandConversionAppService.GetAll();
				List<BrandConversionModel> brandConversionList = brandConversions.DeserializeToBrandConversionList();

				byte[] excelData = ExcelGenerator.ExportBrandConversion(brandConversionList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Brand-Conversions.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Create()
		{
			return PartialView();
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(BrandConversionModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				string data = JsonHelper<BrandConversionModel>.Serialize(model);

				_brandConversionAppService.Add(data);

				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.CreateFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult Delete(long id)
		{
			try
			{
				_brandConversionAppService.Remove(id);

				return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public JsonResult AutoComplete(string prefix)
		{
			ICollection<QueryFilter> filters = new List<QueryFilter>();			
			filters.Add(new QueryFilter("Code", prefix, Operator.Contains));
			filters.Add(new QueryFilter("Description", prefix, Operator.Contains, Operation.Or));
			filters.Add(new QueryFilter("IsActive", "true"));

			string brands = _brandAppService.Find(filters);
			List<BrandModel> brandModelList = brands.DeserializeToBrandList();

			brandModelList = brandModelList.OrderBy(x => x.Code).ToList();			

			return Json(brandModelList, JsonRequestBehavior.AllowGet);
		}

		public ActionResult Edit(long id)
		{
			string BrandConversion = _brandConversionAppService.GetById(id, true);
			BrandConversionModel model = BrandConversion.DeserializeToBrandConversion();

			return PartialView(model);
		}

		[HttpPost]
		public ActionResult Edit(BrandConversionModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidData);
					return RedirectToAction("Index");
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				string data = JsonHelper<BrandConversionModel>.Serialize(model);
				_brandConversionAppService.Update(data);

				SetTrueTempData(UIResources.EditSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.EditFailed);
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		[HttpPost]
		public ActionResult GetAllByLocation()
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
				string brandConversions = _brandConversionAppService.GetAll();
				List<BrandConversionModel> brandConversionList = brandConversions.DeserializeToBrandConversionList();								

				int recordsTotal = brandConversionList.Count();

				// Search    				
				if (!string.IsNullOrEmpty(searchValue))
				{
					brandConversionList = brandConversionList.Where(m => m.BrandCode.ToString().ToLower().Contains(searchValue.ToLower()) ||
												 m.Value1.ToString() == searchValue.ToLower() ||
												 m.Value2.ToString() == searchValue.ToLower() ||
												 m.UOM1 != null && m.UOM1.ToLower().Contains(searchValue.ToLower()) ||
												 m.UOM2 != null && m.UOM2.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "brandcode":
								brandConversionList = brandConversionList.OrderBy(x => x.BrandCode).ToList();
								break;
							case "value1":
								brandConversionList = brandConversionList.OrderBy(x => x.Value1).ToList();
								break;
							case "value2":
								brandConversionList = brandConversionList.OrderBy(x => x.Value1).ToList();
								break;
							case "uom1":
								brandConversionList = brandConversionList.OrderBy(x => x.UOM1).ToList();
								break;
							case "uom2":
								brandConversionList = brandConversionList.OrderBy(x => x.UOM2).ToList();
								break;
							case "notes":
								brandConversionList = brandConversionList.OrderBy(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "brandcode":
								brandConversionList = brandConversionList.OrderByDescending(x => x.BrandCode).ToList();
								break;
							case "value1":
								brandConversionList = brandConversionList.OrderByDescending(x => x.Value1).ToList();
								break;
							case "value2":
								brandConversionList = brandConversionList.OrderByDescending(x => x.Value1).ToList();
								break;
							case "uom1":
								brandConversionList = brandConversionList.OrderByDescending(x => x.UOM1).ToList();
								break;
							case "uom2":
								brandConversionList = brandConversionList.OrderByDescending(x => x.UOM2).ToList();
								break;
							case "notes":
								brandConversionList = brandConversionList.OrderByDescending(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = brandConversionList.Count();

				// Paging     
				var data = brandConversionList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<BrandConversionModel>() }, JsonRequestBehavior.AllowGet);
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
				string BrandConversions = _brandConversionAppService.GetAll();
				List<BrandConversionModel> brandConversionList = BrandConversions.DeserializeToBrandConversionList();

				int recordsTotal = brandConversionList.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					brandConversionList = brandConversionList.Where(m => m.BrandCode.ToString().ToLower().Contains(searchValue.ToLower()) ||
												 m.Value1.ToString() == searchValue.ToLower() ||
												 m.Value2.ToString() == searchValue.ToLower() ||
												 m.UOM1 != null && m.UOM1.ToLower().Contains(searchValue.ToLower()) ||
												 m.UOM2 != null && m.UOM2.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "brandcode":
								brandConversionList = brandConversionList.OrderBy(x => x.BrandCode).ToList();
								break;
							case "value1":
								brandConversionList = brandConversionList.OrderBy(x => x.Value1).ToList();
								break;
							case "value2":
								brandConversionList = brandConversionList.OrderBy(x => x.Value1).ToList();
								break;
							case "uom1":
								brandConversionList = brandConversionList.OrderBy(x => x.UOM1).ToList();
								break;
							case "uom2":
								brandConversionList = brandConversionList.OrderBy(x => x.UOM2).ToList();
								break;
							case "notes":
								brandConversionList = brandConversionList.OrderBy(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "brandcode":
								brandConversionList = brandConversionList.OrderByDescending(x => x.BrandCode).ToList();
								break;
							case "value1":
								brandConversionList = brandConversionList.OrderByDescending(x => x.Value1).ToList();
								break;
							case "value2":
								brandConversionList = brandConversionList.OrderByDescending(x => x.Value1).ToList();
								break;
							case "uom1":
								brandConversionList = brandConversionList.OrderByDescending(x => x.UOM1).ToList();
								break;
							case "uom2":
								brandConversionList = brandConversionList.OrderByDescending(x => x.UOM2).ToList();
								break;
							case "notes":
								brandConversionList = brandConversionList.OrderByDescending(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = brandConversionList.Count();

				// Paging     
				var data = brandConversionList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<BrandConversionModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
