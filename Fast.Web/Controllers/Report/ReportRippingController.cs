using Fast.Application.Interfaces;
using Fast.Web.Models.LPH;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Report
{
    public class ReportRippingController : BaseController<LPHModel>
    {
        private readonly ILPHAppService _lphAppService;
        private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;
        private readonly ILoggerAppService _logger;
        private readonly ILPHApprovalsAppService _lphApprovalAppService;
        private readonly ILPHComponentsAppService _lphComponentsAppService;
        private readonly ILPHLocationsAppService _lphLocationsAppService;
        private readonly ILPHValuesAppService _lphValuesAppService;
        private readonly ILPHValueHistoriesAppService _lphValueHistoriesAppService;
        private readonly ILPHExtrasAppService _lphExtrasAppService;
        private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;

        public ReportRippingController(
          ILPHAppService lphAppService,
          IReferenceAppService referenceAppService,
          ILocationAppService locationAppService,
          ILoggerAppService logger,
            ILPHComponentsAppService lPHComponentsAppService,
            ILPHLocationsAppService lPHLocationsAppService,
            ILPHValuesAppService lPHValuesAppService,
            ILPHApprovalsAppService lPHApprovalsAppService,
            ILPHValueHistoriesAppService lPHValueHistoriesAppService,
            ILPHExtrasAppService lPHExtrasAppService,
            ILPHSubmissionsAppService lPHSubmissionsAppService)
        {
            _lphAppService = lphAppService;
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _logger = logger;
            _lphComponentsAppService = lPHComponentsAppService;
            _lphLocationsAppService = lPHLocationsAppService;
            _lphValuesAppService = lPHValuesAppService;
            _lphApprovalAppService = lPHApprovalsAppService;
            _lphValueHistoriesAppService = lPHValueHistoriesAppService;
            _lphExtrasAppService = lPHExtrasAppService;
            _lphSubmissionsAppService = lPHSubmissionsAppService;
        }
        public ActionResult Index()
        {
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

    }
}
