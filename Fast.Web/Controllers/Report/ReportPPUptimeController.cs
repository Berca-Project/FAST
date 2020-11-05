using Fast.Application.Interfaces;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class ReportPPUptimeController : BaseController<PPLPHModel>
    {
        private readonly IPPLPHAppService _ppLphAppService;
        private readonly IPPLPHApprovalsAppService _ppLphApprovalAppService;
        private readonly IPPLPHComponentsAppService _ppLphComponentsAppService;
        private readonly IPPLPHLocationsAppService _ppLphLocationsAppService;
        private readonly IPPLPHValuesAppService _ppLphValuesAppService;
        private readonly IPPLPHValueHistoriesAppService _ppLphValueHistoriesAppService;
        private readonly IPPLPHExtrasAppService _ppLphExtrasAppService;
        private readonly IPPLPHSubmissionsAppService _ppLphSubmissionsAppService;
        private readonly ILoggerAppService _logger;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IEmployeeAppService _employeeAppService;       
        private readonly IWeeksAppService _weeksAppService;
        private readonly IUserAppService _userAppService;
        private readonly IUserRoleAppService _userRoleAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        public ReportPPUptimeController(
          IPPLPHAppService ppLPHAppService,
          IPPLPHComponentsAppService ppLPHComponentsAppService,
          IPPLPHLocationsAppService ppLPHLocationsAppService,
          IPPLPHValuesAppService ppLPHValuesAppService,
          IPPLPHApprovalsAppService ppLPHApprovalsAppService,
          IPPLPHValueHistoriesAppService ppLPHValueHistoriesAppService,
          IPPLPHExtrasAppService ppLPHExtrasAppService,
          IPPLPHSubmissionsAppService ppLPHSubmissionsAppService,
          ILoggerAppService logger,
          IReferenceAppService referenceAppService,
          ILocationAppService locationAppService,
          IEmployeeAppService employeeAppService,
          IWeeksAppService weeksAppService,
          IUserAppService userAppService,
          IUserRoleAppService userRoleAppService,
          IMachineAppService machineAppService,
          IReferenceDetailAppService referenceDetailAppService)
        {
            _ppLphAppService = ppLPHAppService;
            _ppLphComponentsAppService = ppLPHComponentsAppService;
            _ppLphLocationsAppService = ppLPHLocationsAppService;
            _ppLphValuesAppService = ppLPHValuesAppService;
            _ppLphApprovalAppService = ppLPHApprovalsAppService;
            _ppLphValueHistoriesAppService = ppLPHValueHistoriesAppService;
            _ppLphExtrasAppService = ppLPHExtrasAppService;
            _ppLphSubmissionsAppService = ppLPHSubmissionsAppService;
            _logger = logger;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _employeeAppService = employeeAppService;
            _weeksAppService = weeksAppService;
            _userAppService = userAppService;
            _userRoleAppService = userRoleAppService;
            _machineAppService = machineAppService;
            _referenceDetailAppService = referenceDetailAppService;
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
                string ppList = _ppLphApprovalAppService.GetAll(true);
                List<PPLPHApprovalsModel> pps = ppList.DeserializeToPPLPHApprovalList();
                pps = pps.Where(x => x.Status.Trim().ToLower() == "approved" && x.LocationID == AccountLocationID).ToList();

                foreach (var item in pps)
                {
                    string subs = _ppLphSubmissionsAppService.FindBy("ID", item.LPHSubmissionID, true);
                    PPLPHSubmissionsModel subsModel = subs.DeserializeToPPLPHSubmissions();
                    item.LPHType = subsModel.LPHHeader.Replace("Controller", "");
                }

                pps = pps.Where(x => x.LPHType == ppType).ToList();


                return Json(new { Status = "True" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Result = false;
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { data = new List<PPLPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}
