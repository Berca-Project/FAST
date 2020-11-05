using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
    [CustomAuthorize("accessright")]
    public class AccessRightController : BaseController<AccessRightDBModel>
    {
        #region ::Init::
        private readonly IAccessRightAppService _accessRightAppService;
        private readonly IRoleAppService _roleAppService;
        private readonly IMenuAppService _menuAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;
        #endregion

        #region ::Constructor::
        public AccessRightController(
            IAccessRightAppService accessRightAppService,
            IRoleAppService roleAppService,
            ILoggerAppService logger,
            ILocationAppService locationAppService,
            IReferenceAppService referenceAppService,
            IMenuAppService menuAppService)
        {
            _locationAppService = locationAppService;
            _referenceAppService = referenceAppService;
            _accessRightAppService = accessRightAppService;
            _menuAppService = menuAppService;
            _roleAppService = roleAppService;
            _logger = logger;
        }
        #endregion

        #region ::Public Methods::
        public ActionResult Index(string roleName, long locationID = 0)
        {
            GetTempData();

            ViewBag.RoleList = DropDownHelper.BindDropDownRole(_roleAppService);
            if (AccountIsAdmin)
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            else
                ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService, false, AccountProdCenterID);

            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            ViewBag.SubDepartmentList = DropDownHelper.BuildEmptyList();

            MenuTreeModel menuTree = GetTreeMenu();

            string role = string.IsNullOrEmpty(roleName) ? "SUPERADMIN" : roleName;
            long locID = locationID == 0 ? 1 : locationID;

            AccessRightTreeModel accessRightTree = GetAccessTreeMenu(menuTree, role, locID);
            accessRightTree.Access = GetAccess(WebConstants.MenuSlug.ACCESS_RIGHT, _menuAppService);

            long pcID = 0;
            long depID = 0;
            long subdepID = 0;

            GetPcAndDepID(locationID, out pcID, out depID, out subdepID);

            accessRightTree.ProductionCenterID = pcID;
            accessRightTree.DepartmentID = depID;
            accessRightTree.SubDepartmentID = subdepID;

            if (pcID != 0)
            {
                ViewBag.DepartmentList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, pcID);
            }

            if (depID != 0)
            {
                ViewBag.SubDepartmentList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, depID);
            }

            return View(accessRightTree);
        }

        public ActionResult ExportExcel(long locID)
        {
            try
            {
                // Getting all data    			
                string accessRights = _accessRightAppService.FindBy("LocationID", locID.ToString());
                List<AccessRightDBModel> accessRightList = accessRights.DeserializeToAccessRightList();
                accessRightList = accessRightList.OrderBy(x => x.RoleName).ToList();

                string menus = _menuAppService.GetAll(true);
                List<MenuModel> menuModelList = menus.DeserializeToMenuList();
                List<AccessRightDBModel> result = new List<AccessRightDBModel>();

                foreach (var item in accessRightList)
                {
                    MenuModel menu = menuModelList.Where(x => x.ID == item.MenuID).FirstOrDefault();
                    if (menu != null)
                    {
                        item.MenuName = menu.Name;
                        item.ReadName = item.Read.HasValue ? item.Read.Value ? "Y" : "N" : "N";
                        item.WriteName = item.Write.HasValue ? item.Write.Value ? "Y" : "N" : "N";
                        item.PrintName = item.Print.HasValue ? item.Print.Value ? "Y" : "N" : "N";

                        result.Add(item);
                    }
                }

                string location = _locationAppService.GetLocationFullCode(locID);

                byte[] excelData = ExcelGenerator.ExportMasterAccessRight(result, AccountName, location);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Master-Access-Right.xlsx");
                Response.BinaryWrite(excelData);
                Response.End();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GetDepartmentByProdCenterID(long id)
        {
            List<SelectListItem> _menuList = DropDownHelper.GetDepartmentByProdCenterID(_locationAppService, _referenceAppService, id);

            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetSubDepartmentByDepartmentID(long id)
        {
            List<SelectListItem> _menuList = DropDownHelper.GetSubDepartmentByDepartmentID(_locationAppService, _referenceAppService, id);
            return Json(_menuList, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult UpdatePrivileges(AccessRightLocationModel model)
        {
            try
            {
                string roleName = model.AccessList[0].RoleName;

                ICollection<QueryFilter> filters = new List<QueryFilter>();
                filters.Add(new QueryFilter("RoleName", roleName));
                filters.Add(new QueryFilter("LocationID", model.LocationID.ToString()));
                filters.Add(new QueryFilter("IsDeleted", "0"));

                string ar = _accessRightAppService.FindNoTracking(filters);
                List<AccessRightDBModel> arList = ar.DeserializeToAccessRightList();

                string menus = _menuAppService.GetAll(true);
                List<MenuModel> menuList = menus.DeserializeToMenuList();

                bool isUpdated = false;

                List<long> activeParentIDList = new List<long>();
                // add new privileges
                foreach (var item in model.AccessList)
                {
                    var exist = arList.Where(x => x.MenuID == item.MenuID).FirstOrDefault();
                    if (exist == null)
                    {
                        item.LocationID = model.LocationID;
                        item.RoleName = roleName;
                        item.ModifiedBy = AccountName;
                        item.ModifiedDate = DateTime.Now;

                        string arData = JsonHelper<AccessRightDBModel>.Serialize(item);
                        _accessRightAppService.Add(arData);
                    }
                    else
                    {
                        isUpdated = false;
                        if (exist.Read != item.Read)
                        {
                            isUpdated = true;
                            exist.Read = item.Read;
                        }
                        if (exist.Write != item.Write)
                        {
                            isUpdated = true;
                            exist.Write = item.Write;
                        }
                        if (exist.Print != item.Print)
                        {
                            isUpdated = true;
                            exist.Print = item.Print;
                        }

                        if (isUpdated && !activeParentIDList.Any(x => x == exist.MenuID))
                        {
                            exist.ModifiedBy = AccountName;
                            exist.ModifiedDate = DateTime.Now;

                            string arData = JsonHelper<AccessRightDBModel>.Serialize(exist);
                            _accessRightAppService.Update(arData);
                        }
                    }

                    // check parent menu
                    if ((item.Read.HasValue && item.Read.Value) || (exist != null && exist.Read.HasValue && exist.Read.Value))
                    {
                        var currentMenu = menuList.Where(x => x.ID == item.MenuID).FirstOrDefault();
                        if (currentMenu != null && currentMenu.ParentID.HasValue && currentMenu.ParentID.Value > 0)
                        {
                            var parentExist = arList.Where(x => x.MenuID == currentMenu.ParentID.Value).FirstOrDefault();
                            if (parentExist == null)
                            {
                                AccessRightDBModel newAR = new AccessRightDBModel();
                                newAR.LocationID = model.LocationID;
                                newAR.RoleName = roleName;
                                newAR.ModifiedBy = AccountName;
                                newAR.ModifiedDate = DateTime.Now;
                                newAR.Read = true;

                                string arData = JsonHelper<AccessRightDBModel>.Serialize(newAR);
                                _accessRightAppService.Add(arData);
                            }
                            else if (parentExist.Read.HasValue && !parentExist.Read.Value)
                            {
                                string parent = _accessRightAppService.GetById(parentExist.ID, true);
                                AccessRightDBModel parentModel = parent.DeserializeToAccessRight();

                                parentModel.Read = true;
                                parentModel.ModifiedBy = AccountName;
                                parentModel.ModifiedDate = DateTime.Now;

                                string arData = JsonHelper<AccessRightDBModel>.Serialize(parentModel);
                                _accessRightAppService.Update(arData);
                            }

                            activeParentIDList.Add(currentMenu.ParentID.Value);
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
        #endregion

        #region ::Private Methods::
        private void GetPcAndDepID(long locID, out long prodCenterID, out long deptID, out long subDeptID)
        {
            List<long> result = new List<long>();

            prodCenterID = 0;
            deptID = 0;
            subDeptID = 0;

            LocationModel model = GetLocation(locID);
            if (model.ParentID != 0)
            {
                if (model.ParentID == 1)
                {
                    prodCenterID = model.ID;
                }
                else
                {
                    var model1 = GetLocation(model.ParentID);
                    if (model1.ParentID == 1)
                    {
                        prodCenterID = model1.ID;
                        deptID = model.ID;
                    }
                    else
                    {
                        var model2 = GetLocation(model1.ParentID);
                        if (model2.ParentID == 1)
                        {
                            prodCenterID = model2.ID;
                            deptID = model1.ID;
                            subDeptID = model.ID;
                        }
                    }
                }
            }
        }

        private LocationModel GetLocation(long locationID)
        {
            string location = _locationAppService.GetById(locationID, true);
            LocationModel locationModel = location.DeserializeToLocation();

            return locationModel;
        }

        private AccessRightTreeModel GetAccessTreeMenu(MenuTreeModel menuTree, string roleName, long locationID)
        {
            AccessRightTreeModel accessRightTree = new AccessRightTreeModel();
            accessRightTree.RoleName = roleName;

            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("RoleName", roleName));
            filters.Add(new QueryFilter("LocationID", locationID.ToString()));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string accessRights = _accessRightAppService.FindNoTracking(filters);
            List<AccessRightDBModel> accessRightList = accessRights.DeserializeToAccessRightList();

            foreach (var parentMenu in menuTree.Parents)
            {
                ParentAccessRightModel parentModel = new ParentAccessRightModel();
                parentModel.MenuID = parentMenu.ID;
                parentModel.MenuName = parentMenu.Name;

                if (parentMenu.IsTopMenu)
                {
                    AccessRightDBModel currentAR = accessRightList.Where(x => x.MenuID == parentMenu.ID).FirstOrDefault();
                    parentModel.ID = currentAR == null ? 0 : currentAR.ID;
                    parentModel.RoleName = roleName;
                    parentModel.Read = currentAR == null ? false : currentAR.Read;
                    parentModel.Write = currentAR == null ? false : currentAR.Write;
                    parentModel.Print = currentAR == null ? false : currentAR.Print;

                    accessRightTree.Parents.Add(parentModel);
                }
                else
                {
                    bool isAllRead = true;
                    bool isAllWrite = true;
                    bool isAllPrint = true;

                    foreach (var childMenu in parentMenu.Children)
                    {
                        ChildAccessRightModel childModel = new ChildAccessRightModel();
                        childModel.MenuID = childMenu.ID;
                        childModel.MenuName = childMenu.Name;

                        AccessRightDBModel currentAR = accessRightList.Where(x => x.MenuID == childMenu.ID).FirstOrDefault();
                        childModel.ID = currentAR == null ? 0 : currentAR.ID;
                        childModel.RoleName = roleName;
                        childModel.Read = currentAR == null ? false : currentAR.Read;
                        childModel.Write = currentAR == null ? false : currentAR.Write;
                        childModel.Print = currentAR == null ? false : currentAR.Print;
                        childModel.Print = currentAR == null ? false : currentAR.Print;

                        if (childModel.Read == null || !childModel.Read.Value)
                            isAllRead = false;
                        if (childModel.Write == null || !childModel.Write.Value)
                            isAllWrite = false;
                        if (childModel.Print == null || !childModel.Print.Value)
                            isAllPrint = false;

                        parentModel.Children.Add(childModel);
                    }

                    parentModel.Read = isAllRead;
                    parentModel.Write = isAllWrite;
                    parentModel.Print = isAllPrint;

                    accessRightTree.Parents.Add(parentModel);
                }
            }

            return accessRightTree;
        }

        private MenuTreeModel GetTreeMenu()
        {
            MenuTreeModel model = new MenuTreeModel();

            // get all menu
            string allmenu = _menuAppService.GetAll(true);
            List<ChildMenuModel> allMenuList = allmenu.DeserializeToChildMenuList();
            List<ParentMenuModel> parentList = allmenu.DeserializeToParentMenuList().Where(x => x.IsParent).ToList();
            List<ParentMenuModel> topmenuList = allmenu.DeserializeToParentMenuList().Where(x => x.IsTopMenu).ToList();

            List<ParentMenuModel> parentListResult = new List<ParentMenuModel>();

            // populate child menu list
            foreach (var parent in parentList.ToList())
            {
                var childList = allMenuList.Where(x => x.ParentID == parent.ID).OrderBy(x => x.DisplayOrder).ToList();

                ParentMenuModel parentTemp = Copy(parent);

                parentTemp.Children.AddRange(childList);

                parentListResult.Add(parentTemp);
            }

            // add parent list
            model.Parents.AddRange(parentListResult);

            // add top menu list
            model.Parents.AddRange(topmenuList);

            model.Parents = model.Parents.OrderBy(x => x.DisplayOrder).ToList();

            return model;
        }

        private ParentMenuModel Copy(ParentMenuModel menu)
        {
            ParentMenuModel result = new ParentMenuModel
            {
                Access = menu.Access,
                Description = menu.Description,
                DisplayOrder = menu.DisplayOrder,
                ID = menu.ID,
                IsActive = menu.IsActive,
                IsDeleted = menu.IsDeleted,
                IsParent = menu.IsParent,
                IsTopMenu = menu.IsTopMenu,
                ModifiedBy = menu.ModifiedBy,
                ModifiedDate = menu.ModifiedDate,
                Name = menu.Name,
                PageAction = menu.PageAction,
                PageController = menu.PageController,
                PageIcon = menu.PageIcon,
                PageSlug = menu.PageSlug,
                ParentID = menu.ParentID,
                ParentName = menu.ParentName
            };

            return result;
        }

        private RoleModel GetRole(string roleName)
        {
            string role = _roleAppService.GetByName(roleName, true);

            return role.DeserializeToRole();
        }

        private MenuModel GetMenu(long menuId)
        {
            string menu = _menuAppService.GetById(menuId, true);

            return menu.DeserializeToMenu();
        }

        private AccessRightDBModel GetAccessRight(long accessRightID)
        {
            string accessRight = _accessRightAppService.GetById(accessRightID, true);

            return accessRight.DeserializeToAccessRight();
        }
        #endregion
    }
}
