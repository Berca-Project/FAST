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
    public class ReportLPHandLPPCController : BaseController<LPHModel>
    {
        private readonly ILPHAppService _lphAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;
        private readonly IBrandAppService _brandAppService;
        private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
        private readonly ILPHApprovalsAppService _lphApprovalAppService;

        public ReportLPHandLPPCController(
         ILPHAppService lphAppService,
          IReferenceAppService referenceAppService,
          ILocationAppService locationAppService,
          ILoggerAppService logger,
          IBrandAppService brandAppService,
          ILPHSubmissionsAppService lPHSubmissionsAppService,
          ILPHApprovalsAppService lPHApprovalsAppService)
        {
            _lphAppService = lphAppService;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _logger = logger;
            _brandAppService = brandAppService;
            _lphSubmissionsAppService = lPHSubmissionsAppService;
            _lphApprovalAppService = lPHApprovalsAppService;
        }
     
        public ActionResult Index()
        {
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);
            ViewBag.DepartmentList = DropDownHelper.BuildEmptyList();
            return View();
        }


        [HttpPost]
        public ActionResult GetReportWithParam(string spType, string startDate, string endDate, string week, long productionCenterID) //param nya menyesuaikan
        {
            try
            {
                string spList = _lphApprovalAppService.GetAll(true);
                List<LPHApprovalsModel> sp = spList.DeserializeToLPHApprovalList();
                sp = sp.Where(x => x.Status.Trim().ToLower() == "approved" && x.LocationID == AccountLocationID).ToList();

                foreach (var item in sp)
                {
                    string subs = _lphSubmissionsAppService.FindBy("ID", item.LPHSubmissionID, true);
                    LPHSubmissionsModel subsModel = subs.DeserializeToLPHSubmissions();
                    item.LPHType = subsModel.LPHHeader.Replace("Controller", "");
                }

                sp = sp.Where(x => x.LPHType == spType).ToList();

                if (startDate != "" && endDate != "")
                {
                    DateTime startDt = DateTime.Parse(startDate);
                    DateTime endDt = DateTime.Parse(endDate);
                }

                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<LPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GenerateExcel(DateTime dtFilter, long prodCenterID)
        {
            try
            {
                //byte[] excelData = ExcelGenerator.ExportLPHandLPPC(AccountName);

                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("content-disposition", "attachment;filename=LPH and LPPC.xlsx");
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
