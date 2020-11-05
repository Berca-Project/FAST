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
    [CustomAuthorize("menu")]
    public class MenuController : BaseController<MenuModel>
    {
		#region ::Init::
		private readonly IMenuAppService _menuAppService;
        private readonly ILoggerAppService _logger;
		#endregion

		#region ::Constructor::
		public MenuController(IMenuAppService menuAppService, ILoggerAppService logger)
        {
            _menuAppService = menuAppService;
            _logger = logger;
        }
		#endregion

		#region ::Public Methods::
		public ActionResult Index()
        {
			GetTempData();

            MenuTreeModel model = GetIndexModel();

            return View(model);
        }

		public static List<SelectListItem> BindDropDownPageIcon()
		{
			List<SelectListItem> _menuList = new List<SelectListItem>();

			#region
			_menuList.Add(new SelectListItem
			{
				Text = "File",
				Value = "file"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Menu",
				Value = "menu"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Layers",
				Value = "layers"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "User",
				Value = "user"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Users",
				Value = "users"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Pie Chart",
				Value = "pie-chart"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Box",
				Value = "box"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "CPU",
				Value = "cpu"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Grid",
				Value = "grid"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Save",
				Value = "save"
			});

			_menuList.Add(new SelectListItem
			{
				Text = "Share",
				Value = "share-2"
			});

			_menuList = _menuList.OrderBy(x => x.Text).ToList();
			#endregion

			return _menuList;
		}
        
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(MenuTreeModel menuModel)
        {
            try
            {
                if (!ModelState.IsValid)
                {
					SetFalseTempData(UIResources.InvalidModelState);
                    return RedirectToAction("Index");
                }

                string exist = _menuAppService.GetBy("Name", menuModel.Name);
                if (!string.IsNullOrEmpty(exist))
                {
					SetFalseTempData(string.Format(UIResources.DataExist, "Menu", menuModel.Name));
					return RedirectToAction("Index");
				}

                if (menuModel.IsParent || menuModel.IsTopMenu)
                {
                    menuModel.ParentID = 0;
                }
                else
                {
                    if (string.IsNullOrEmpty(menuModel.PageController) || string.IsNullOrEmpty(menuModel.PageAction))
                    {
						SetFalseTempData(UIResources.PageControllerAndActionRequired);
						return RedirectToAction("Index");
					}
                }

                menuModel.IsActive = true;
                menuModel.ModifiedBy = AccountName;
                menuModel.ModifiedDate = DateTime.Now;

                string data = JsonHelper<MenuModel>.Serialize(menuModel);

                _menuAppService.Add(data);

				SetTrueTempData(UIResources.CreateSucceed);
            }
            catch (Exception ex)
            {
				SetFalseTempData(UIResources.CreateFailed);				
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

			return RedirectToAction("Index");
		}
        
        public ActionResult Edit(int id)
        {
            ViewBag.ParentMenuList = DropDownHelper.BindDropDownParentMenu(_menuAppService);
            ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
			ViewBag.IconList = BindDropDownPageIcon();

			MenuModel model = GetMenu(id);

            return PartialView(model);
        }
        
        [HttpPost]
        public ActionResult Edit(MenuModel menuModel)
        {
            try
            {
                if (!ModelState.IsValid)
                {
					SetFalseTempData(UIResources.InvalidModelState);
					return RedirectToAction("Index");
				}

                if (menuModel.IsParent || menuModel.IsTopMenu)
                {
                    menuModel.ParentID = 0;
                }
                else
                {
                    if (string.IsNullOrEmpty(menuModel.PageController) || string.IsNullOrEmpty(menuModel.PageAction))
                    {
						SetFalseTempData(UIResources.PageControllerAndActionRequired);
						return RedirectToAction("Index");
					}
                }

                menuModel.ModifiedBy = AccountName;
                menuModel.ModifiedDate = DateTime.Now;

                string data = JsonHelper<MenuModel>.Serialize(menuModel);

                _menuAppService.Update(data);

				SetTrueTempData(UIResources.EditSucceed);
            }
            catch (Exception ex)
            {
				SetFalseTempData(UIResources.EditFailed);				

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

			return RedirectToAction("Index");
		}

        public ActionResult Delete(int id)
        {
            MenuModel menu = GetMenu(id);

            return PartialView(menu);
        }

        [HttpPost]
        public ActionResult Delete(long id)
        {
            try
            {
                MenuModel Menu = GetMenu(id);
                Menu.IsDeleted = true;

                string MenuData = JsonHelper<MenuModel>.Serialize(Menu);
                _menuAppService.Update(MenuData);

				SetTrueTempData(UIResources.DeleteSucceed);
            }
            catch (Exception ex)
            {
				SetTrueTempData(UIResources.DeleteSucceed);

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
            }

            return RedirectToAction("Index");
        }

		public ActionResult ExportExcel()
		{
			try
			{
				// Getting all data    			
				string menuList = _menuAppService.GetAll(true);
				List<MenuModel> menuModelList = menuList.DeserializeToMenuList();
				List<MenuModel> results = new List<MenuModel>();
				List<MenuModel> parentList = menuModelList.Where(x => x.IsParent || x.IsTopMenu).OrderBy(x => x.DisplayOrder).ToList();

				// populate all parent and top menu
				foreach (var item in parentList)
				{
					if (item.IsTopMenu)
					{
						item.ParentName = item.Name;
						item.Name = "-";

						results.Add(item);
					}
					else
					{
						// get all children
						List<MenuModel> children = menuModelList.Where(x => x.ParentID == item.ID).OrderBy(x => x.DisplayOrder).ToList();
						foreach (var child in children)
						{
							child.ParentName = item.Name;
							results.Add(child);
						}
					}
				}

				byte[] excelData = ExcelGenerator.ExportMasterMenu(results, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Menu.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}
		#endregion

		#region ::Private Methods::
		private MenuTreeModel GetIndexModel()
		{
			ViewBag.ParentMenuList = DropDownHelper.BindDropDownParentMenu(_menuAppService);
			ViewBag.UserStatusList = DropDownHelper.BindDropDownUserStatus();
			ViewBag.IconList = BindDropDownPageIcon();

			MenuTreeModel model = new MenuTreeModel();
			model.IsActive = true;
			model.Access = GetAccess(WebConstants.MenuSlug.MENU, _menuAppService);

			// get all menu
			string allmenu = _menuAppService.GetAll(true);
			List<ChildMenuModel> allMenuList = allmenu.DeserializeToChildMenuList();
			List<ParentMenuModel> parentList = allmenu.DeserializeToParentMenuList().Where(x => x.IsParent).ToList();
			List<ParentMenuModel> topmenuList = allmenu.DeserializeToParentMenuList().Where(x => x.IsTopMenu).ToList();

			// populate child menu list
			foreach (var parent in parentList)
			{
				var childList = allMenuList.Where(x => x.ParentID == parent.ID).OrderBy(x => x.DisplayOrder).ToList();
				parent.Children.AddRange(childList);
			}

			// add parent list
			model.Parents.AddRange(parentList);

			// add top menu list
			model.Parents.AddRange(topmenuList);

			model.Parents = model.Parents.OrderBy(x => x.DisplayOrder).ToList();

			return model;
		}

		private MenuModel GetMenu(long menuId)
        {
            string menu = _menuAppService.GetById(menuId, true);
            MenuModel menuModel = menu.DeserializeToMenu();

            return menuModel;
        }
		#endregion
	}
}
