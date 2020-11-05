using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace Fast.Web.Controllers.LPH
{

    public class TimbanganController : BaseController<AccessRightDBModel>
    {
        private readonly ILPHAppService _lphAppService;
        private readonly ILPHComponentsAppService _lphComponentsAppService;
        private readonly ILPHLocationsAppService _lphLocationsAppService;
        private readonly ILPHValuesAppService _lphValuesAppService;
        private readonly ILPHApprovalsAppService _lphApprovalAppService;
        private readonly ILPHValueHistoriesAppService _lphValueHistoriesAppService;
        private readonly ILPHExtrasAppService _lphExtrasAppService;
        private readonly ILoggerAppService _logger;
        private readonly IReferenceAppService _referenceAppService;
        private readonly IMppAppService _mppAppService;
        private readonly IWppAppService _wppAppService;
        private readonly IWeeksAppService _weeksAppService;
        private readonly IUserAppService _userAppService;
        private readonly IReferenceDetailAppService _referenceDetailAppService;
        private readonly ILPHSubmissionsAppService _lphSubmissionsAppService;
        private readonly IMachineAppService _machineAppService;
        private readonly IEmployeeAppService _employeeAppService;
        private readonly IBrandAppService _brandAppService;
        private readonly ILocationAppService _locationAppService;

        public TimbanganController(
            ILPHAppService lphAppService,
			IBrandAppService brandAppService,
            IWppAppService wppAppService,
            IWeeksAppService weeksAppService,
            ILPHComponentsAppService lphComponentsAppService,
            ILPHLocationsAppService lphLocationsAppService,
            ILPHValuesAppService lphValuesAppService,
            ILPHApprovalsAppService lphApprovalsAppService,
            ILPHValueHistoriesAppService lphValueHistoriesAppService,
            ILPHExtrasAppService lphExtrasAppService,
            ILoggerAppService logger,
            IReferenceAppService referenceAppService,
            IUserAppService userAppService,
            IMppAppService mppAppService,
            IReferenceDetailAppService referenceDetailAppService,
            IMachineAppService machineAppService,
            IEmployeeAppService employeeAppService,
             ILocationAppService locationAppService,
              ILPHSubmissionsAppService lPHSubmissionsAppService)
        {
            _lphAppService = lphAppService;
            _locationAppService = locationAppService;
            _lphComponentsAppService = lphComponentsAppService;
            _lphLocationsAppService = lphLocationsAppService;
            _lphValuesAppService = lphValuesAppService;
            _lphApprovalAppService = lphApprovalsAppService;
            _lphValueHistoriesAppService = lphValueHistoriesAppService;
            _lphExtrasAppService = lphExtrasAppService;
            _logger = logger;
            _referenceAppService = referenceAppService;
            _userAppService = userAppService;
            _mppAppService = mppAppService;
            _referenceDetailAppService = referenceDetailAppService;
            _weeksAppService = weeksAppService;
            _wppAppService = wppAppService;
            _lphSubmissionsAppService = lPHSubmissionsAppService;
            _machineAppService = machineAppService;
            _employeeAppService = employeeAppService;
            _brandAppService = brandAppService;
        }
        [HttpPost]
        public ActionResult GetBeratSpec(string brand)
        {
            try
            {
                List<ReferenceDetailModel> referenceResult = new List<ReferenceDetailModel>();
                string getIDReference = _referenceAppService.GetBy("Name", "Berat Cigarette", true);
                ReferenceModel referenceData = getIDReference.DeserializeToReference();
                if (referenceData != null)
                {
                    string resultCloveCon = _referenceDetailAppService.FindBy("ReferenceID", referenceData.ID, true);
                    List<ReferenceDetailModel> cloveConList = resultCloveCon.DeserializeToRefDetailList();
                    referenceResult = cloveConList.Where(x => x.Code == brand).ToList();
                }

				if (referenceResult.Count > 0)
					return Json(new { Status = "True", Data = referenceResult[0].Description }, JsonRequestBehavior.AllowGet);
				else
					return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
            catch (Exception ex)
            {
                _logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
                return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
            }
        }
        // GET: Timbangan
        public ActionResult Index()
        {
            return View();
        }

        // GET: Timbangan/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }
    }
}
