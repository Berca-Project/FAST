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
	[CustomAuthorize("userlogs")]
	public class LogsController : BaseController<UserLogModel>
	{
		private readonly IUserLogAppService _userLogAppService;
		private readonly ILoggerAppService _logger;
		private readonly IMenuAppService _menuService;

		public LogsController(IUserLogAppService userLogService, IMenuAppService menuService, ILoggerAppService logger)
		{
			_userLogAppService = userLogService;
			_menuService = menuService;
			_logger = logger;
		}

		// GET: UserLog
		public ActionResult Index()
		{
			UserLogModel model = new UserLogModel();
			model.Access = GetAccess(WebConstants.MenuSlug.USER_LOGS, _menuService);

			return View(model);
		}
		
		public ActionResult Delete(UserLogModel model)
		{
			try
			{
				ICollection<QueryFilter> filters = new List<QueryFilter>();
				filters.Add(new QueryFilter("TimeStamp", model.DeletedDate.Value.Date.AddDays(1).ToString(), Operator.LessThan));

				string logs = _userLogAppService.Find(filters);
				if (!string.IsNullOrEmpty(logs))
				{
					_userLogAppService.RemoveRange(logs);

					ViewBag.Result = true;					
				}
				else
				{
					ViewBag.Result = false;
					ViewBag.ErrorMessage = UIResources.NoLogs;
				}
			}
			catch(Exception ex)
			{
				ViewBag.Result = false;
				ViewBag.ErrorMessage = UIResources.DeleteFailed;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			UserLogModel usermodel = new UserLogModel();
			usermodel.Access = GetAccess(WebConstants.MenuSlug.USER_LOGS, _menuService);

			return View("Index", usermodel);
		}

		public ActionResult ExportExcel()
		{
			try
			{
				// Getting all data    			
				string logs = _userLogAppService.GetAll();
				List<UserLogModel> logList = logs.DeserializeToUserLogList();
				logList = logList.OrderByDescending(x => x.Timestamp).ToList();

				foreach (var item in logList)
				{
					item.Message = item.Message.Length > 200 ? item.Message.Substring(0, 200) : item.Message;
				}

				byte[] excelData = ExcelGenerator.ExportMasterUserLogs(logList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Master-UserLogs.xlsx");
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
				string logs = _userLogAppService.GetAll();
				List<UserLogModel> logList = logs.DeserializeToUserLogList();
				logList = logList.OrderByDescending(x => x.Timestamp).ToList();

				int recordsTotal = logList.Count();

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					logList = logList.Where(m => m.UserID.ToString().ToLower().Contains(searchValue.ToLower()) ||
                                                 m.Level.ToLower().Contains(searchValue.ToLower()) ||
                                                 m.Stacktrace.ToLower().Contains(searchValue.ToLower()) ||                                                 
                                                 m.Message.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "userid":
								logList = logList.OrderBy(x => x.UserID).ToList();
								break;
							case "timestamp":
								logList = logList.OrderBy(x => x.Timestamp).ToList();
								break;
                            case "level":
                                logList = logList.OrderBy(x => x.Level).ToList();
                                break;
                            default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "userid":
								logList = logList.OrderByDescending(x => x.UserID).ToList();
								break;
							case "timestamp":
								logList = logList.OrderByDescending(x => x.Timestamp).ToList();
								break;
                            case "level":
                                logList = logList.OrderByDescending(x => x.Level).ToList();
                                break;
                            default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = logList.Count();

				// Paging     
				var data = logList.Skip(skip).Take(pageSize).ToList();

				foreach (var item in data)
				{
					item.Message = item.Message.Length > 200 ? item.Message.Substring(0, 200) : item.Message;
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<UserLogModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
