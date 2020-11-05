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
	[CustomAuthorize("machine")]
	public class MachineAllocationController : BaseController<ReferenceDetailModel>
	{
		private readonly IReferenceAppService _referenceAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;
		private readonly IMachineAppService _machineAppService;
		private readonly IMachineAllocationAppService _machineAllocationService;
        private readonly IRoleAppService _roleAppService;

        public MachineAllocationController(
			IReferenceAppService referenceAppService,
            IRoleAppService roleAppService,
            IMenuAppService menuService,
			IMachineAppService machineAppService,
			IMachineAllocationAppService locMachineAllocationService,
			ILoggerAppService logger)
		{
			_referenceAppService = referenceAppService;
			_menuService = menuService;
			_logger = logger;
            _roleAppService = roleAppService;
            _machineAllocationService = locMachineAllocationService;
			_machineAppService = machineAppService;
		}

		// GET: MachineAllocation
		public ActionResult Index()
		{
			GetTempData();

			//ViewBag.MachineCategoryList = DropDownHelper.BuildEmptyList();
            ViewBag.MachineCategoryList = DropDownHelper.BindDropDownRole(_roleAppService); //DropDownHelper.BindDropDownMachineCategory(_referenceAppService);

            MachineAllocationModel model = new MachineAllocationModel();
			model.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

			return View(model);
		}

		[HttpPost]
		public JsonResult AutoComplete(string prefix)
		{
			ICollection<QueryFilter> filters = new List<QueryFilter>();
			filters.Add(new QueryFilter("Code", prefix, Operator.Contains));
			filters.Add(new QueryFilter("IsDeleted", "0", Operator.Equals));

			string machineList = _machineAppService.Find(filters);
			List<MachineModel> machineModelList = machineList.DeserializeToMachineList();

			machineModelList = machineModelList.OrderBy(x => x.Code).ToList();

			return Json(machineModelList, JsonRequestBehavior.AllowGet);
		}

		// GET: MachineAllocation/Create
		public ActionResult Create()
		{
			//ViewBag.MachineCategoryList = DropDownHelper.BindDropDownMachineCategory(_referenceAppService);
            ViewBag.MachineCategoryList = DropDownHelper.BindDropDownRole(_roleAppService); 

            return PartialView();
		}

		// POST: MachineAllocation/Create
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(MachineAllocationModel machineAllocationModel)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (machineAllocationModel.MachineID == 0)
				{
					SetFalseTempData(UIResources.MachineNotSelected);
					return RedirectToAction("Index");
				}

				if (string.IsNullOrEmpty(machineAllocationModel.MachineCategory))
				{
					SetFalseTempData("No Machine category is selected");
					return RedirectToAction("Index");
				}

                if (machineAllocationModel.Value == 0)
                {
                    SetFalseTempData("Value is missing");
                    return RedirectToAction("Index");
                }

                string machine = _machineAppService.GetById(machineAllocationModel.MachineID);
				MachineModel machineModel = machine.DeserializeToMachine();

				machineAllocationModel.MachineCode = machineModel.Code;
				machineAllocationModel.ModifiedDate = DateTime.Now;
				machineAllocationModel.ModifiedBy = AccountName;

				string data = JsonHelper<MachineAllocationModel>.Serialize(machineAllocationModel);

				_machineAllocationService.Add(data);


				SetTrueTempData(UIResources.CreateSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.CreateFailed);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}


		// GET: MachineAllocation/Edit/5
		public ActionResult Edit(int id)
		{
			MachineAllocationModel model = GetMachineAllocation(id);

			ViewBag.MachineCategoryList = DropDownHelper.BindDropDownRole(_roleAppService); //DropDownHelper.BindDropDownMachineCategory(_referenceAppService);

            return PartialView(model);
		}

		// POST: MachineAllocation/Edit/5
		[HttpPost]
		public ActionResult Edit(MachineAllocationModel machineAllocationModel)
		{
			try
			{
				machineAllocationModel.Access = GetAccess(WebConstants.MenuSlug.MACHINE, _menuService);

				if (!ModelState.IsValid)
				{
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

				if (machineAllocationModel.MachineID == 0)
				{
					SetFalseTempData(UIResources.MachineNotSelected);
					return RedirectToAction("Index");
				}

                if (machineAllocationModel.Value == 0)
                {
                    SetFalseTempData("Value is missing");
                    return RedirectToAction("Index");
                }

                if (string.IsNullOrEmpty(machineAllocationModel.MachineCategory))
				{
					SetFalseTempData("No Machine category is selected");
					return RedirectToAction("Index");
				}

				string machine = _machineAppService.GetById(machineAllocationModel.MachineID);
				MachineModel machineModel = machine.DeserializeToMachine();

				machineAllocationModel.MachineCode = machineModel.Code;
				machineAllocationModel.ModifiedBy = AccountName;
				machineAllocationModel.ModifiedDate = DateTime.Now;

				string data = JsonHelper<MachineAllocationModel>.Serialize(machineAllocationModel);

				_machineAllocationService.Update(data);

				SetTrueTempData(UIResources.EditSucceed);
			}
			catch (Exception ex)
			{
				SetFalseTempData(UIResources.EditFailed);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		// POST: MachineAllocation/Delete/5
		[HttpPost]
		public ActionResult Delete(long id)
		{
			try
			{
				//MachineAllocationModel machineAllocation = GetMachineAllocation(id);
				//machineAllocation.IsDeleted = true;

				//string userData = JsonHelper<MachineAllocationModel>.Serialize(machineAllocation);
				_machineAllocationService.Remove(id);

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
				string machineAllocationList = _machineAllocationService.GetAll(true);
				List<MachineAllocationModel> machineAllocationModelList = machineAllocationList.DeserializeToMachineAllocationList();

				int recordsTotal = machineAllocationModelList.Count();				

				sortColumn = sortColumn == "ID" ? "" : sortColumn;

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					machineAllocationModelList = machineAllocationModelList.Where(m => m.MachineCode.ToLower().Contains(searchValue.ToLower()) ||
													 m.MachineCategory.ToLower().Contains(searchValue.ToLower())).ToList();

				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "machinecode":
								machineAllocationModelList = machineAllocationModelList.OrderBy(x => x.MachineCode).ToList();
								break;
							case "machinecategory":
								machineAllocationModelList = machineAllocationModelList.OrderBy(x => x.MachineCategory).ToList();
								break;
                            case "value":
                                machineAllocationModelList = machineAllocationModelList.OrderBy(x => x.Value).ToList();
                                break;
                            default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "machinecode":
								machineAllocationModelList = machineAllocationModelList.OrderByDescending(x => x.MachineCode).ToList();
								break;
							case "machinecategory":
								machineAllocationModelList = machineAllocationModelList.OrderByDescending(x => x.MachineCategory).ToList();
								break;
                            case "value":
                                machineAllocationModelList = machineAllocationModelList.OrderByDescending(x => x.Value).ToList();
                                break;
                            default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = machineAllocationModelList.Count();

				// Paging     
				var data = machineAllocationModelList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<MachineAllocationModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		private MachineAllocationModel GetMachineAllocation(long machineAllocationID)
		{
			string machineAllocation = _machineAllocationService.GetById(machineAllocationID, true);
			MachineAllocationModel model = machineAllocation.DeserializeToMachineAllocation();

			return model;
		}

        public ActionResult ExportExcel()
        {
            try
            {
                // Getting all data    			
                string machineAllocationList = _machineAllocationService.GetAll(true);
                List<MachineAllocationModel> machineAllocationModelList = machineAllocationList.DeserializeToMachineAllocationList();

                byte[] excelData = ExcelGenerator.ExportMachineAllocation(machineAllocationModelList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-Machine-Allocation.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }
    }
}
