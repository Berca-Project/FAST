using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Checklist
{
	[CustomAuthorize("generator")]
	public class GeneratorController : BaseController<AccessRightDBModel>
	{
		private readonly IChecklistAppService _checklistAppService;
		private readonly IChecklistLocationAppService _checklistLocationAppService;
		private readonly IChecklistComponentAppService _checklistComponentAppService;
		private readonly IChecklistValueAppService _checklistValueAppService;
		private readonly ILoggerAppService _logger;
		private readonly IReferenceAppService _referenceAppService;

		public GeneratorController(
			IChecklistAppService checklistAppService,
			IChecklistLocationAppService checklistLocationAppService,
			IChecklistComponentAppService checklistComponentAppService,
			IChecklistValueAppService checklistValueAppService,
			ILoggerAppService logger,
			IReferenceAppService referenceAppService)
		{
			_checklistAppService = checklistAppService;
			_checklistLocationAppService = checklistLocationAppService;
			_checklistComponentAppService = checklistComponentAppService;
			_checklistValueAppService = checklistValueAppService;
			_logger = logger;
			_referenceAppService = referenceAppService;
		}

		private List<SelectListItem> BindDropDownLocation()
		{
			string locationPC = _referenceAppService.GetDetailAll(ReferenceEnum.ProdCenter, true);
			List<ReferenceDetailModel> locationPCList = string.IsNullOrEmpty(locationPC) ? new List<ReferenceDetailModel>() : JsonConvert.DeserializeObject<List<ReferenceDetailModel>>(locationPC);
			List<SelectListItem> _menuList = new List<SelectListItem>();
			locationPCList = locationPCList.OrderBy(x => x.Description).ToList();

			foreach (var location in locationPCList)
			{
				_menuList.Add(new SelectListItem
				{
					Text = location.Description,
					Value = location.ID.ToString()
				});
			}

			return _menuList;
		}

		// GET: Generator
		public ActionResult Index()
		{
            string checklists = _checklistAppService.GetAll(true);
            IEnumerable<ChecklistModel> checklistsList = string.IsNullOrEmpty(checklists) ? new List<ChecklistModel>() : JsonConvert.DeserializeObject<IEnumerable<ChecklistModel>>(checklists);

            ViewBag.Checklists = checklistsList;

            return View();
		}

        /*
        // GET: Generator/Details/5
        public ActionResult Details(int id)
		{
			return View();
		}


		// GET: Generator/Create
		public ActionResult Create()
		{
			ViewBag.LocationList = BindDropDownLocation();

			return View();
		}

		// POST: Generator/Create
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Create(ChecklistModel model)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					return RedirectToAction("Create");
				}

				model.ModifiedBy = AccountName;
				model.ModifiedDate = DateTime.Now;

				foreach (var locat in model.LocationIDs)
				{
					var modelLocation = new ChecklistLocationModel();
					modelLocation.LocationID = locat;
					model.Locations.Add(modelLocation);
				}

				string data = JsonHelper<ChecklistModel>.Serialize(model);
				_checklistAppService.Add(data);

				string checklists = _checklistAppService.GetAll();
				IEnumerable<ChecklistModel> checklistsList = string.IsNullOrEmpty(checklists) ? new List<ChecklistModel>() : JsonConvert.DeserializeObject<IEnumerable<ChecklistModel>>(checklists);

				var result = checklistsList.Where(y => y.ModifiedBy == AccountName && !y.IsDeleted).OrderByDescending(x => x.ID).FirstOrDefault();

				return RedirectToAction("Content/" + result.ID);
			}
			catch (Exception ex)
			{
				ViewBag.LocationList = BindDropDownLocation();

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return View();
			}
		}

		// GET: Generator/Content
		public ActionResult Content(long ID)
		{
			string checklist = _checklistAppService.GetById(ID);
			ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : JsonConvert.DeserializeObject<ChecklistModel>(checklist);
			ViewBag.Checklist = currentChecklist;

            string location_id = _checklistLocationAppService.FindBy("ChecklistID", currentChecklist.ID.ToString(), true);
            var location_ids = string.IsNullOrEmpty(location_id) ? new List<long>() : JsonConvert.DeserializeObject<List<ChecklistLocationModel>>(location_id).Select(x=>x.LocationID).ToList();

            string locationPC = _referenceAppService.GetDetailAll(ReferenceEnum.ProdCenter, true);
            List<ReferenceDetailModel> locationPCList = string.IsNullOrEmpty(locationPC) ? new List<ReferenceDetailModel>() : JsonConvert.DeserializeObject<List<ReferenceDetailModel>>(locationPC);
            ViewBag.Locations = locationPCList.Where(x => location_ids.Contains(x.ID)).Select(x=>x.Description).ToList();

            return View();
		}

		// POST: Generator/Content
		[HttpPost]
		public ActionResult Content(long ID, string result)
		{
			try
			{
				result = result.Replace("[[", "[").Replace("]]", "]");

				List<SortableResultModel> results = string.IsNullOrEmpty(result) ? new List<SortableResultModel>() : JsonConvert.DeserializeObject<List<SortableResultModel>>(result);

				SaveComponents(ID, results, 0);

				return RedirectToAction("ContentDetail/" + ID);
			}
			catch (Exception ex)
			{
				string checklist = _checklistAppService.GetById(ID);
				ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : JsonConvert.DeserializeObject<ChecklistModel>(checklist);
				ViewBag.Checklist = currentChecklist;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return View();
			}
		}

		// GET: Generator/ContentDetail
		public ActionResult ContentDetail(long ID)
		{
			string checklist = _checklistAppService.GetById(ID);
			ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : JsonConvert.DeserializeObject<ChecklistModel>(checklist);
			ViewBag.Checklist = currentChecklist;

            string location_id = _checklistLocationAppService.FindBy("ChecklistID", currentChecklist.ID.ToString(), true);
            var location_ids = string.IsNullOrEmpty(location_id) ? new List<long>() : JsonConvert.DeserializeObject<List<ChecklistLocationModel>>(location_id).Select(x => x.LocationID).ToList();

            string locationPC = _referenceAppService.GetDetailAll(ReferenceEnum.ProdCenter, true);
            List<ReferenceDetailModel> locationPCList = string.IsNullOrEmpty(locationPC) ? new List<ReferenceDetailModel>() : JsonConvert.DeserializeObject<List<ReferenceDetailModel>>(locationPC);
            ViewBag.Locations = locationPCList.Where(x => location_ids.Contains(x.ID)).Select(x => x.Description).ToList();

            string component = _checklistComponentAppService.FindByNoTracking("ChecklistID", currentChecklist.ID.ToString(), true);
			List<ChecklistComponentModel> components = string.IsNullOrEmpty(component) ? new List<ChecklistComponentModel>() : JsonConvert.DeserializeObject<List<ChecklistComponentModel>>(component);

			return View(components);
		}

		// POST: Generator/Content
		[HttpPost]
		public ActionResult ContentDetail(long ID, List<ChecklistComponentModel> results)
		{
            if (results.Count() > 0)
            {
                foreach(var result in results)
                {
                    var temp_data = new ChecklistComponentDBModel();

                    Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

                    if (result.ComponentFile != null && result.ComponentFile.ContentLength > 0)
                    {
                        result.ComponentName = unixTimestamp.ToString() + "_" + result.ComponentFile.FileName;
                        result.ComponentFile.SaveAs(Server.MapPath("~/upload/") + result.ComponentName);
                    }

                    temp_data.ID = result.ID;
                    temp_data.ChecklistID = result.ChecklistID;
                    temp_data.ComponentType = result.ComponentType;
                    temp_data.ComponentName = result.ComponentName;
                    temp_data.AdditionalClass = result.ComponentType;
                    temp_data.OrderNum = result.OrderNum;
                    temp_data.Parent = result.Parent;

                    string data = JsonHelper<ChecklistComponentDBModel>.Serialize(temp_data);

                    _checklistComponentAppService.Update(data);
                }
            }

            return RedirectToAction("Generated/" + ID);
        }

        // GET: Generator/ContentDetail
        public ActionResult Generated(long ID)
        {
            string checklist = _checklistAppService.GetById(ID);
            ChecklistModel currentChecklist = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : JsonConvert.DeserializeObject<ChecklistModel>(checklist);
            ViewBag.Checklist = currentChecklist;

            string component = _checklistComponentAppService.FindBy("ChecklistID", currentChecklist.ID.ToString(), true);
            List<ChecklistComponentModel> components = string.IsNullOrEmpty(component) ? new List<ChecklistComponentModel>() : JsonConvert.DeserializeObject<List<ChecklistComponentModel>>(component);

            return View(components);
        }

        // POST: Generator/Content
        [HttpPost]
        public ActionResult Generated(long ID, List<ChecklistValueModel> results)
        {
            if (results.Count() > 0)
            {
                foreach (var result in results)
                {
                    result.Shift = results[0].Shift;
                    result.date = results[0].date;
                    result.UserID = AccountID;
                    result.ValueDate = DateTime.Now;

                    string data = JsonHelper<ChecklistValueModel>.Serialize(result);
                    try { _checklistValueAppService.Add(data); }
                    catch(Exception e)
                    {

                    }
                   
                }
            }

            return RedirectToAction("Generated/" + ID);
        }

        // GET: Generator/Delete/5
        public ActionResult Delete(long id)
		{
            try
            {
                string checklist = _checklistAppService.GetById(id, true);
                ChecklistModel model = string.IsNullOrEmpty(checklist) ? new ChecklistModel() : JsonConvert.DeserializeObject<ChecklistModel>(checklist);

                model.IsDeleted = true;

                string data = JsonHelper<ChecklistModel>.Serialize(model);
                _checklistAppService.Update(data);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
		}

		// POST: Generator/Delete/5
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
		public void SaveComponents(long ChecklistID, List<SortableResultModel> m, long parent)
		{
			try
			{				
				ChecklistComponentModel result = new ChecklistComponentModel();

				PopulateChild(ChecklistID, m, result, parent);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}
		}

		private void PopulateChild(long ChecklistID, List<SortableResultModel> nodes, ChecklistComponentModel result, long parent)
		{
			int counter = 1;
			foreach (var child in nodes)
			{
				var model = new ChecklistComponentModel();
				model.ChecklistID = ChecklistID;
				model.ComponentType = child.id.Trim();				
				model.OrderNum = counter++;
				model.Parent = parent;
				model.ModifiedBy = AccountName;
				result.Children.Add(model);

				string data = JsonHelper<ChecklistComponentModel>.Serialize(model);

				long id =_checklistComponentAppService.Add(data);

				if (child.children != null && child.children.Count > 0)
				{
					PopulateChild(ChecklistID, child.children, model, id);
				}
			}
		}
        */

	}
}
