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
    [CustomAuthorize("role")]
    public class RoleController : BaseController<JobTitleModel>
    {
        private readonly IJobTitleAppService _jobTitleAppService;
        private readonly IRoleAppService _roleAppService;
        private readonly ILoggerAppService _logger;
        private readonly IMenuAppService _menuService;

        public RoleController(
            IJobTitleAppService jobTitleAppService,
            ILoggerAppService logger,
            IMenuAppService menuService,
            IRoleAppService roleAppService)
        {
            _jobTitleAppService = jobTitleAppService;
            _roleAppService = roleAppService;
            _logger = logger;
            _menuService = menuService;
        }

        // GET: UserType
        public ActionResult Index()
        {
            GetTempData();

            JobTitleTreeModel model = GetIndexModel();

            return View(model);
        }

        private JobTitleTreeModel GetIndexModel()
        {
            ViewBag.RoleList = DropDownHelper.BuildEmptyList();

            JobTitleTreeModel model = GetTreeJobTitle();

            return model;
        }

        public ActionResult ExportExcel()
        {
            try
            {
                // Getting all data    			
                string jobTitleList = _jobTitleAppService.GetAll(true);
                List<JobTitleModel> jobTitles = jobTitleList.DeserializeToJobTitleList();
                var unassignTitles = jobTitles.Where(x => string.IsNullOrEmpty(x.RoleName)).ToList();
                jobTitles = jobTitles.Where(x => !string.IsNullOrEmpty(x.RoleName)).OrderBy(x => x.RoleName).ToList();
                jobTitles.AddRange(unassignTitles);

                byte[] excelData = ExcelGenerator.ExportMasterRole(jobTitles, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-Roles.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        private JobTitleTreeModel GetTreeJobTitle()
        {
            JobTitleTreeModel model = new JobTitleTreeModel();
            List<ParentJobTitleModel> parentModelList = new List<ParentJobTitleModel>();

            // get parent list			
            string roles = _roleAppService.GetAll(true);
            List<RoleModel> parentList = roles.DeserializeToRoleList();
            foreach (var parent in parentList)
            {
                ParentJobTitleModel tempParent = new ParentJobTitleModel
                {
                    Description = parent.Description,
                    Name = parent.Name
                };

                parentModelList.Add(tempParent);
            }

            // get all job title
            string jobTitles = _jobTitleAppService.GetAll(true);
            List<JobTitleModel> jobTitleList = jobTitles.DeserializeToJobTitleList();

            // populate child menu list
            foreach (var parent in parentModelList)
            {
                var childList = jobTitleList.Where(x => x.RoleName == parent.Name).OrderBy(x => x.Title).ToList();

                parent.Children.AddRange(childList);
            }

            // add parent list
            model.Parents.AddRange(parentModelList);

            model.Parents = model.Parents.OrderBy(x => x.Name).ToList();

            model.Access = GetAccess(WebConstants.MenuSlug.ROLE, _menuService);

            return model;
        }

        public ActionResult Edit(int id)
        {
            JobTitleModel jobTitle = GetJobTitle(id);

            ViewBag.RoleList = DropDownHelper.BuildEmptyList();

            return PartialView(jobTitle);
        }

        [HttpPost]
        public ActionResult Edit(JobTitleModel jtModel)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                jtModel.ModifiedBy = AccountName;
                jtModel.ModifiedDate = DateTime.Now;

                string data = JsonHelper<JobTitleModel>.Serialize(jtModel);

                _jobTitleAppService.Update(data);

                SetTrueTempData(UIResources.EditSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.EditFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult EditMaster(string name)
        {
            ViewBag.RoleList = DropDownHelper.BuildEmptyList();

            string role = _roleAppService.GetByName(name, true);
            RoleModel roleModel = role.DeserializeToRole();

            return PartialView(roleModel);
        }

        [HttpPost]
        public ActionResult EditMaster(RoleModel roleModel)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                roleModel.ModifiedBy = AccountName;
                roleModel.ModifiedDate = DateTime.Now;

                string data = JsonHelper<RoleModel>.Serialize(roleModel);

                _roleAppService.Update(data);

                SetTrueTempData(UIResources.EditSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(UIResources.EditFailed);
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        public ActionResult Manage()
        {
            List<SelectListItem> temp = DropDownHelper.BindDropDownRoleAddNew(_roleAppService);
            ViewBag.RoleList = temp;
            if (temp.Count > 1)
            {
                ViewBag.AvailableTitles = DropDownHelper.BindDropDownJobTitleUnAssigned(_jobTitleAppService);
                ViewBag.SelectedTitles = DropDownHelper.BindDropDownJobTitleByRole(_jobTitleAppService, temp[1].Text);
            }
            else
            {
                ViewBag.AvailableTitles = DropDownHelper.BindDropDownJobTitleUnAssigned(_jobTitleAppService);
                ViewBag.SelectedTitles = DropDownHelper.BuildEmpty();
            }

            JobTitleTreeModel model = new JobTitleTreeModel();
            model.Access = GetAccess(WebConstants.MenuSlug.ROLE, _menuService);

            return PartialView(model);
        }

        [HttpPost]
        public ActionResult Manage(string name, string description, string rolename, List<string> selectedTitles)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                if (rolename == "0")
                {
                    RoleModel newRole = new RoleModel();
                    newRole.Name = name;
                    newRole.Description = description;
                    newRole.ModifiedBy = AccountName;
                    newRole.ModifiedDate = DateTime.Now;

                    string newRoleEntity = JsonHelper<RoleModel>.Serialize(newRole);

                    _roleAppService.Add(newRoleEntity);

                    rolename = name;
                }

                string role = _roleAppService.GetByName(rolename, true);
                RoleModel roleModel = role.DeserializeToRole();

                string jobTitle = _jobTitleAppService.FindByNoTracking("IsDeleted", "0", true);
                List<JobTitleModel> jobTitleList = jobTitle.DeserializeToJobTitleList();
                jobTitleList = jobTitleList.Where(x => x.RoleName == rolename).ToList();

                // check any new title
                if (selectedTitles != null)
                {
                    foreach (var item in selectedTitles)
                    {
                        if (!jobTitleList.Any(x => x.Title == item))
                        {
                            string newTitle = _jobTitleAppService.FindByNoTracking("Title", item, true);
                            JobTitleModel newTitleModel = newTitle.DeserializeToJobTitleList().FirstOrDefault();
                            if (newTitleModel != null)
                            {
                                newTitleModel.RoleName = rolename;
                                newTitleModel.ModifiedBy = AccountName;
                                newTitleModel.ModifiedDate = DateTime.Now;

                                string updatemodel = JsonHelper<JobTitleModel>.Serialize(newTitleModel);

                                _jobTitleAppService.Update(updatemodel);
                            }
                        }
                    }
                }

                // check any removed title
                foreach (var item in jobTitleList)
                {
                    if (selectedTitles == null || !selectedTitles.Any(x => x == item.Title))
                    {
                        item.RoleName = null;
                        item.ModifiedBy = AccountName;
                        item.ModifiedDate = DateTime.Now;

                        string removedItem = JsonHelper<JobTitleModel>.Serialize(item);

                        _jobTitleAppService.Update(removedItem);
                    }
                }

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult AddTitle()
        {
            ViewBag.RoleList = DropDownHelper.BindDropDownRole(_roleAppService);

            return PartialView();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult AddTitle(JobTitleModel model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return RedirectToAction("Index");
                }

                string exist = _jobTitleAppService.GetBy("Title", model.Title);
                if (!string.IsNullOrEmpty(exist))
                {
                    SetFalseTempData(string.Format(UIResources.DataExist, "JobTitle", model.Title));

                    return RedirectToAction("Index");
                }

                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;

                string data = JsonHelper<JobTitleModel>.Serialize(model);

                _jobTitleAppService.Add(data);

                SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(JobTitleTreeModel model)
        {
            try
            {
                List<SelectListItem> temp = DropDownHelper.BindDropDownRole(_roleAppService);

                ViewBag.RoleList = temp;
                ViewBag.JobTitleList = DropDownHelper.BindDropDownJobTitle(_jobTitleAppService, temp[0].Value);

                if (!ModelState.IsValid)
                {
                    SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                string exist = _jobTitleAppService.GetBy("Code", model.Code);
                JobTitleModel jobTitleModel = exist.DeserializeToJobTitle();
                if (!string.IsNullOrEmpty(exist) && jobTitleModel.RoleName.Equals(model.RoleName))
                {
                    SetFalseTempData(string.Format(UIResources.DataExist, "JobTitle", model.Title));
                    return RedirectToAction("Index");
                }

                model.ModifiedBy = AccountName;
                model.ModifiedDate = DateTime.Now;

                string data = JsonHelper<JobTitleModel>.Serialize(model);

                _jobTitleAppService.Add(data);

                SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
                SetFalseTempData(ex.Message);

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult Delete(long id)
        {
            try
            {
                _jobTitleAppService.Remove(id);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult DeleteRole(string name)
        {
            try
            {
                _roleAppService.RemoveEntity(name);

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GetJobTitleByRole(string role)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string jobs = _jobTitleAppService.FindBy("RoleName", role, true);
            List<JobTitleModel> jobList = jobs.DeserializeToJobTitleList();
            foreach (var item in jobList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Title,
                    Value = item.ID.ToString()
                });
            }

            List<SelectListItem> titleList = DropDownHelper.BindDropDownJobTitleUnAssigned(_jobTitleAppService);
            foreach (var item in titleList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Text,
                    Value = "#" + item.Value
                });
            }

            return Json(_menuList, JsonRequestBehavior.AllowGet);
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
                string jobTitleList = _jobTitleAppService.GetAll(true);
                List<JobTitleModel> jobTitles = jobTitleList.DeserializeToJobTitleList();
                jobTitles = jobTitles.OrderBy(x => x.RoleName).ToList();

                int recordsTotal = jobTitles.Count();

                // total number of rows count     
                int recordsFiltered = jobTitles.Count();

                // Paging     
                var data = jobTitles.Skip(skip).Take(pageSize).ToList();

                // Returning Json Data    
                return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;

                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

                return Json(new { data = new List<JobTitleModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        private JobTitleModel GetJobTitle(long jobTitleID)
        {
            string jobTitle = _jobTitleAppService.GetById(jobTitleID, true);
            JobTitleModel model = jobTitle.DeserializeToJobTitle();

            return model;
        }
    }
}