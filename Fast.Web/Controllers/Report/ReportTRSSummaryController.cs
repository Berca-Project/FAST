using Fast.Application.Interfaces;
using Fast.Web.Models.LPH;
using Fast.Web.Resources;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class ReportTRSSummaryController : BaseController<LPHModel>
    {
        private readonly ILPHAppService _lphAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;
        public ReportTRSSummaryController(
           ILPHAppService lphAppService,
           IReferenceAppService referenceAppService,
           ILocationAppService locationAppService,
           ILoggerAppService logger)
            {
                _lphAppService = lphAppService;
                _referenceAppService = referenceAppService;
                _locationAppService = locationAppService;
                _logger = logger;
        }
     
        public ActionResult Index()
        {
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            return View();
        }
        
        [HttpPost]
        public ActionResult GetReportByParam(string startDate, string endDate, string week,  long locationID)
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

                if (startDate != "" && endDate !="")
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

        [HttpPost]
        public ActionResult GenerateExcel(DateTime dtFilter, long prodCenterID)
        {
            try
            {
                //byte[] excelData = ExcelGenerator.ExportTRSSummary(AccountName);

                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("content-disposition", "attachment;filename=TRSSummary.xlsx");
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
