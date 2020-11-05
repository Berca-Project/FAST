using Fast.Application.Interfaces;
using Fast.Web.Models;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class ReportPPSavetyDeviceController : BaseController<PPLPHModel>
    {
        private readonly IChecklistAppService _checklistAppService;
        private readonly IChecklistLocationAppService _checklistLocationAppService;
        private readonly IChecklistComponentAppService _checklistComponentAppService;
        private readonly IChecklistSubmitAppService _checklistSubmitAppService;
        private readonly IChecklistValueAppService _checklistValueAppService;
        private readonly IChecklistValueHistoryAppService _checklistValueHistoryAppService;
        private readonly IChecklistApprovalAppService _checklistApprovalAppService;
        private readonly IChecklistApproverAppService _checklistApproverAppService;
        private readonly ILoggerAppService _logger;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IEmployeeAppService _employeeAppService;
        private readonly IWeeksAppService _weeksAppService;       
      
      
        public ReportPPSavetyDeviceController(
        IChecklistAppService checklistAppService,
        IChecklistLocationAppService checklistLocationAppService,
        IChecklistComponentAppService checklistComponentAppService,
        IChecklistSubmitAppService checklistSubmitAppService,
        IChecklistValueAppService checklistValueAppService,
        IChecklistValueHistoryAppService checklistValueHistoryAppService,
        IChecklistApprovalAppService checklistApprovalAppService,
        IChecklistApproverAppService checklistApproverAppService,
         ILoggerAppService logger,
         IReferenceAppService referenceAppService,
         ILocationAppService locationAppService,
         IEmployeeAppService employeeAppService,
         IWeeksAppService weeksAppService
      )
        {
            _checklistAppService = checklistAppService;
            _checklistLocationAppService = checklistLocationAppService;
            _checklistComponentAppService = checklistComponentAppService;
            _checklistSubmitAppService = checklistSubmitAppService;
            _checklistValueAppService = checklistValueAppService;
            _checklistValueHistoryAppService = checklistValueHistoryAppService;
            _checklistApprovalAppService = checklistApprovalAppService;
            _checklistApproverAppService = checklistApproverAppService;
            _logger = logger;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _employeeAppService = employeeAppService;
            _weeksAppService = weeksAppService;         
        
        }
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GetReportWithParam(string ppType) //param nya menyusul
        {
            try
            {
                


                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<ChecklistSubmitModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}
