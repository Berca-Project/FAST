using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.LPH.Primary
{
	[CustomAuthorize("approvalprimary")]
	public class ApprovalPrimaryController : BaseController<PPLPHApprovalsModel>
	{
		private readonly IPPLPHAppService _ppLphAppService;
		private readonly IPPLPHApprovalsAppService _ppLphApprovalAppService;
		private readonly IPPLPHSubmissionsAppService _ppLphSubmissionsAppService;
		private readonly ILoggerAppService _logger;
		private readonly IEmployeeAppService _employeeAppService;
		private readonly IUserAppService _userAppService;
		private readonly ILocationAppService _locationAppService;
        private readonly IReferenceAppService _referenceAppService;


        public ApprovalPrimaryController(
		  IPPLPHAppService ppLphAppService,
		  IPPLPHApprovalsAppService ppLphApprovalsAppService,
		  IPPLPHSubmissionsAppService ppLphSubmissionsAppService,
		  ILoggerAppService logger,
          IReferenceAppService referenceAppService,
          IEmployeeAppService employeeAppService,
		  ILocationAppService locationAppService,
		  IUserAppService userAppService)
		{
			_logger = logger;
			_ppLphAppService = ppLphAppService;
			_ppLphApprovalAppService = ppLphApprovalsAppService;
			_ppLphSubmissionsAppService = ppLphSubmissionsAppService;
			_employeeAppService = employeeAppService;
            _referenceAppService = referenceAppService;
            _userAppService = userAppService;
			_locationAppService = locationAppService;
		}
		// GET: ApprovalPrimary
		public ActionResult Index()
		{
            ViewBag.LPHTypeList = BindDropDownPPLPHType().OrderBy(x => x.Text).ToList();
            LocationTreeModel LocationTree = GetLocationTreeModel();
            ViewBag.LocationTree = LocationTree;
            return View();
        }
		#region Get Data

		[HttpPost]
        public ActionResult GetData(string dateFilter = "", string shift = "", string lphtype = "", string status = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
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

				// Getting all data approvals   			
				//string approvals = _lphApprovalAppService.FindBy("LocationID", AccountLocationID.ToString(), true);
				string approvals = _ppLphApprovalAppService.GetAll();
				List<PPLPHApprovalsModel> approvalList = approvals.DeserializeToPPLPHApprovalList().OrderByDescending(x => x.Date).ToList();

                // get user id list 
                //List<long> userIdList = GetUserIDList(AccountEmployeeID);
                long currentUserID = AccountID;
                approvalList = approvalList.OrderByDescending(x => x.ID).ToList();

                // Getting all data lph               
                string lphList = _ppLphAppService.GetAll(true);
                List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();

                // Getting all data lphsubmission               
                string lphSubList = _ppLphSubmissionsAppService.GetAll(true);
                List<PPLPHSubmissionsModel> lphSubs = lphSubList.DeserializeToPPLPHSubmissionsList();

                //chanif: exclude LPH yg sudah dihapus
                foreach (var item in lphSubs.ToList())
                {
                    var check = lphs.Where(x => x.ID == item.LPHID).FirstOrDefault();
                    // chanif: exclude LPH yang sudah dihapus
                    if (check == null)
                    {
                        lphSubs.Remove(item);
                        continue;
                    }
                    else if (check.IsDeleted)
                    {
                        lphSubs.Remove(item);
                        continue;
                    }
                }

                Dictionary<long, string> locationMap = new Dictionary<long, string>();

                if (!string.IsNullOrEmpty(main_source))
                {
                    List<long> locations = new List<long>();

                    if (main_source == "MyLocat")
                    {
                        //langsung anggap aja punya sub
                        locations.Add(AccountLocationID);

                        string deps = _locationAppService.FindBy("ParentID", AccountLocationID, true);
                        var depsM = deps.DeserializeToLocationList();

                        foreach (var dep in depsM)
                        {
                            locations.Add(dep.ID);

                            string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                            var subdepsM = subdeps.DeserializeToLocationList();

                            foreach (var subdep in subdepsM)
                            {
                                locations.Add(subdep.ID);
                            }
                        }

                        lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
                        //submissions = submissions.Where(x => x.LocationID == AccountLocationID).ToList();
                    }
                    else if (main_source == "Location")
                    {
                        if (!string.IsNullOrEmpty(location3))
                        {
                            lphSubs = lphSubs.Where(x => x.LocationID == Int64.Parse(location3)).ToList();
                        }
                        else if (!string.IsNullOrEmpty(location2))
                        {
                            locations.Add(Int64.Parse(location2));

                            string subdeps = _locationAppService.FindBy("ParentID", location2, true);
                            var subdepsM = subdeps.DeserializeToLocationList();

                            foreach (var subdep in subdepsM)
                            {
                                locations.Add(subdep.ID);
                            }

                            lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
                        }
                        else if (!string.IsNullOrEmpty(location1))
                        {
                            locations.Add(Int64.Parse(location1));

                            string deps = _locationAppService.FindBy("ParentID", location1, true);
                            var depsM = deps.DeserializeToLocationList();

                            foreach (var dep in depsM)
                            {
                                locations.Add(dep.ID);

                                string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                                var subdepsM = subdeps.DeserializeToLocationList();

                                foreach (var subdep in subdepsM)
                                {
                                    locations.Add(subdep.ID);
                                }
                            }

                            lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
                        }
                        //kalau kosong semua ya skip
                    }
                    //jika getall abaikan filtering
                }
                else //location saya
                {
                    approvalList = approvalList.Where(x => x.ApproverID == currentUserID).ToList();
                    // di atas sudah ada kondisi jika approvernya saya
                }
                if (!string.IsNullOrEmpty(dateFilter))
                {
                    DateTime dateFL = DateTime.Parse(dateFilter);
                    approvalList = approvalList.Where(x => x.Date == dateFL.Date).ToList();
                }
                if (!string.IsNullOrEmpty(shift))
                {
                    lphSubs = lphSubs.Where(x => x.Shift.Trim() == shift).ToList();
                }
                if (!string.IsNullOrEmpty(lphtype))
                {
                    lphtype = lphtype.Trim();

                    if (lphtype == "Kretek Line - Addback")
                        lphtype = "LPHPrimaryKretekLineAddback";
                    else if (lphtype == "Intermediate Line - DIET")
                        lphtype = "LPHPrimaryDiet";
                    else if (lphtype == "Intermediate Line - Clove Feeding & DCCC")
                        lphtype = "LPHPrimaryCloveInfeedConditioning";
                    else if (lphtype == "Intermediate Line - CSF Cut Dry & Packing")
                        lphtype = "LPHPrimaryCSFCutDryPacking";
                    else if (lphtype == "Intermediate Line - CSF Feeding & DCCC")
                        lphtype = "LPHPrimaryCSFInfeedConditioning";
                    else if (lphtype == "Intermediate Line - Clove Cut Dry & Packing")
                        lphtype = "LPHPrimaryCloveCutDryPacking";
                    else if (lphtype == "Intermediate Line - RTC")
                        lphtype = "LPHPrimaryRTC";
                    else if (lphtype == "Intermediate Line - Casing Kitchen")
                        lphtype = "LPHPrimaryKitchen";
                    else if (lphtype == "White Line OTP - Process Note")
                        lphtype = "LPHPrimaryWhiteLineOTP";
                    else if (lphtype == "Kretek Line - Feeding KR & RJ")
                        lphtype = "LPHPrimaryKretekLineFeeding";
                    else if (lphtype == "Kretek Line - DCCC KR & RJ")
                        lphtype = "LPHPrimaryKretekLineConditioning";
                    else if (lphtype == "Kretek Line - Cut Dry")
                        lphtype = "LPHPrimaryKretekLineCuttingDrying";
                    else if (lphtype == "Kretek Line - Packing")
                        lphtype = "LPHPrimaryKretekLinePacking";
                    else if (lphtype == "Kretek Line - CRES Feeding & DCCC")
                        lphtype = "LPHPrimaryCresFeedingConditioning";
                    else if (lphtype == "Kretek Line - CRES Cut Dry & Packing")
                        lphtype = "LPHPrimaryCresDryingPacking";
                    else if (lphtype == "White Line PMID - Feeding White")
                        lphtype = "LPHPrimaryWhiteLineFeedingWhite";
                    else if (lphtype == "White Line PMID - DCCC")
                        lphtype = "LPHPrimaryWhiteLineDCCC";
                    else if (lphtype == "White Line PMID - Cutting + FTD")
                        lphtype = "LPHPrimaryWhiteLineCuttingFTD";
                    else if (lphtype == "White Line PMID - Addback")
                        lphtype = "LPHPrimaryWhiteLineAddback";
                    else if (lphtype == "White Line PMID - Packing White")
                        lphtype = "LPHPrimaryWhiteLinePackingWhite";
                    else if (lphtype == "White Line PMID - Feeding SPM")
                        lphtype = "LPHPrimaryWhiteLineFeedingSPM";
                    else if (lphtype == "White Line PMID - Feeding IS White")
                        lphtype = "LPHPrimaryISWhiteFeeding";
                    else if (lphtype == "White Line PMID - Cut Dry IS White")
                        lphtype = "LPHPrimaryISWhiteCutDry";

                    lphSubs = lphSubs.Where(x => x.LPHHeader == lphtype + "Controller").ToList();
                }


                List<long> sudah = new List<long>();
                foreach (var item in approvalList.ToList())
                {
                    if (sudah.Contains(item.LPHSubmissionID))
                    {
                        approvalList.Remove(item); // chanif: hanya munculkan status approval terakhir
                        continue;
                    }

                    if (item.Status.Trim().ToLower() == "submitted" || item.Status.Trim() == "")
                    {
                        item.Status = "Waiting for Approval";
                        if (item.ApproverID == currentUserID || item.UserID == currentUserID)
                            item.IsNeedMyApproval = true;
                    }

                    var lphsub = lphSubs.Where(x => x.ID == item.LPHSubmissionID).FirstOrDefault();
                    if (lphsub == null)
                    {
                        approvalList.Remove(item); // chanif: jika tidak ada datanya berarti draft, jangan munculkan approval
                        continue;
                    }
                    else
                    {
                        item.Location = "-";
                        item.User = "-";

                        if (lphsub.Location != null)
                            item.Location = lphsub.Location;
                        if (lphsub.UserFullName != null)
                            item.User = lphsub.UserFullName;
                    }

                    var check = lphs.Where(x => x.ID == lphsub.LPHID).FirstOrDefault();
                    // chanif: exclude LPH yang sudah dihapus
                    if (check == null)
                    {
                        approvalList.Remove(item);
                        continue;
                    }
                    else if (check.IsDeleted)
                    {
                        approvalList.Remove(item);
                        continue;
                    }

                    var temp = check.MenuTitle;
                    item.LPHType = temp.Replace("Controller", "");

                    sudah.Add(item.LPHSubmissionID);
                }

                if (!string.IsNullOrEmpty(status))
                {
                    if (status.Trim().ToLower() == "approved")
                        approvalList = approvalList.Where(x => x.Status.Trim().ToLower() == "approved").ToList();
                }
                else
                {
                    approvalList = approvalList.Where(x => x.Status == "Waiting for Approval").ToList();
                }
                int recordsTotal = approvalList.Count();

				// Search    - Correction 231019
				if (!string.IsNullOrEmpty(searchValue))
				{
					approvalList = approvalList.Where(m => (m.Shift != null ? m.Shift.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Location != null ? m.Location.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.LPHType != null ? m.LPHType.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.User != null ? m.User.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Approver != null ? m.Approver.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Date != null ? m.Date.ToString("dd-MMM-yy").ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.LPHSubmissionID != 0 ? m.LPHSubmissionID.ToString().Contains(searchValue.ToLower()) : false) ||
											(m.Notes != null ? m.Notes.ToLower().Contains(searchValue.ToLower()) : false) ||
											(m.Status != null ? m.Status.ToLower().Contains(searchValue.ToLower()) : false)).ToList();

				}

				if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDir)))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "lphsubmissionid":
								approvalList = approvalList.OrderBy(x => x.LPHSubmissionID).ToList();
								break;
							case "location":
								approvalList = approvalList.OrderBy(x => x.Location).ToList();
								break;
							case "date":
								approvalList = approvalList.OrderBy(x => x.Date).ToList();
								break;
							case "lphtype":
								approvalList = approvalList.OrderBy(x => x.LPHType).ToList();
								break;
							case "user":
								approvalList = approvalList.OrderBy(x => x.User).ToList();
								break;
							case "approver":
								approvalList = approvalList.OrderBy(x => x.Approver).ToList();
								break;
							case "status":
								approvalList = approvalList.OrderBy(x => x.Status).ToList();
								break;
							case "notes":
								approvalList = approvalList.OrderBy(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "lphsubmissionid":
								approvalList = approvalList.OrderByDescending(x => x.LPHSubmissionID).ToList();
								break;
							case "location":
								approvalList = approvalList.OrderByDescending(x => x.Location).ToList();
								break;
							case "date":
								approvalList = approvalList.OrderByDescending(x => x.Date).ToList();
								break;
							case "lphtype":
								approvalList = approvalList.OrderByDescending(x => x.LPHType).ToList();
								break;
							case "user":
								approvalList = approvalList.OrderByDescending(x => x.User).ToList();
								break;
							case "approver":
								approvalList = approvalList.OrderByDescending(x => x.Approver).ToList();
								break;
							case "status":
								approvalList = approvalList.OrderByDescending(x => x.Status).ToList();
								break;
							case "notes":
								approvalList = approvalList.OrderByDescending(x => x.Notes).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = approvalList.Count();

				// Paging     
				var data = approvalList.Skip(skip).Take(pageSize).ToList();

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{

				var aa = ex.InnerException;

				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<PPLPHSubmissionsModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

        public ActionResult ExportExcel(string dateFilter = "", string shift = "", string lphtype = "", string status = "", string main_source = "", string location1 = "", string location2 = "", string location3 = "")
        {
            try
            {
                // Getting all data approvals   			
                //string approvals = _lphApprovalAppService.FindBy("LocationID", AccountLocationID.ToString(), true);
                string approvals = _ppLphApprovalAppService.GetAll();
                List<PPLPHApprovalsModel> approvalList = approvals.DeserializeToPPLPHApprovalList().OrderByDescending(x => x.Date).ToList();

                // get user id list 
                //List<long> userIdList = GetUserIDList(AccountEmployeeID);
                long currentUserID = AccountID;
                approvalList = approvalList.OrderByDescending(x => x.ID).ToList();

                // Getting all data lph               
                string lphList = _ppLphAppService.GetAll(true);
                List<PPLPHModel> lphs = lphList.DeserializeToPPLPHList();

                // Getting all data lphsubmission               
                string lphSubList = _ppLphSubmissionsAppService.GetAll(true);
                List<PPLPHSubmissionsModel> lphSubs = lphSubList.DeserializeToPPLPHSubmissionsList();

                //chanif: exclude LPH yg sudah dihapus
                foreach (var item in lphSubs.ToList())
                {
                    var check = lphs.Where(x => x.ID == item.LPHID).FirstOrDefault();
                    // chanif: exclude LPH yang sudah dihapus
                    if (check == null)
                    {
                        lphSubs.Remove(item);
                        continue;
                    }
                    else if (check.IsDeleted)
                    {
                        lphSubs.Remove(item);
                        continue;
                    }
                }

                Dictionary<long, string> locationMap = new Dictionary<long, string>();

                if (!string.IsNullOrEmpty(main_source))
                {
                    List<long> locations = new List<long>();

                    if (main_source == "MyLocat")
                    {
                        //langsung anggap aja punya sub
                        locations.Add(AccountLocationID);

                        string deps = _locationAppService.FindBy("ParentID", AccountLocationID, true);
                        var depsM = deps.DeserializeToLocationList();

                        foreach (var dep in depsM)
                        {
                            locations.Add(dep.ID);

                            string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                            var subdepsM = subdeps.DeserializeToLocationList();

                            foreach (var subdep in subdepsM)
                            {
                                locations.Add(subdep.ID);
                            }
                        }

                        lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
                        //submissions = submissions.Where(x => x.LocationID == AccountLocationID).ToList();
                    }
                    else if (main_source == "Location")
                    {
                        if (!string.IsNullOrEmpty(location3))
                        {
                            lphSubs = lphSubs.Where(x => x.LocationID == Int64.Parse(location3)).ToList();
                        }
                        else if (!string.IsNullOrEmpty(location2))
                        {
                            locations.Add(Int64.Parse(location2));

                            string subdeps = _locationAppService.FindBy("ParentID", location2, true);
                            var subdepsM = subdeps.DeserializeToLocationList();

                            foreach (var subdep in subdepsM)
                            {
                                locations.Add(subdep.ID);
                            }

                            lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
                        }
                        else if (!string.IsNullOrEmpty(location1))
                        {
                            locations.Add(Int64.Parse(location1));

                            string deps = _locationAppService.FindBy("ParentID", location1, true);
                            var depsM = deps.DeserializeToLocationList();

                            foreach (var dep in depsM)
                            {
                                locations.Add(dep.ID);

                                string subdeps = _locationAppService.FindBy("ParentID", dep.ID, true);
                                var subdepsM = subdeps.DeserializeToLocationList();

                                foreach (var subdep in subdepsM)
                                {
                                    locations.Add(subdep.ID);
                                }
                            }

                            lphSubs = lphSubs.Where(x => locations.Contains(x.LocationID)).ToList();
                        }
                        //kalau kosong semua ya skip
                    }
                    //jika getall abaikan filtering
                }
                else //location saya
                {
                    approvalList = approvalList.Where(x => x.ApproverID == currentUserID).ToList();
                    // di atas sudah ada kondisi jika approvernya saya
                }
                if (!string.IsNullOrEmpty(dateFilter))
                {
                    DateTime dateFL = DateTime.Parse(dateFilter);
                    approvalList = approvalList.Where(x => x.Date == dateFL.Date).ToList();
                }
                if (!string.IsNullOrEmpty(shift))
                {
                    lphSubs = lphSubs.Where(x => x.Shift.Trim() == shift).ToList();
                }
                if (!string.IsNullOrEmpty(lphtype))
                {
                    lphtype = lphtype.Trim();

                    if (lphtype == "Kretek Line - Addback")
                        lphtype = "LPHPrimaryKretekLineAddback";
                    else if (lphtype == "Intermediate Line - DIET")
                        lphtype = "LPHPrimaryDiet";
                    else if (lphtype == "Intermediate Line - Clove Feeding & DCCC")
                        lphtype = "LPHPrimaryCloveInfeedConditioning";
                    else if (lphtype == "Intermediate Line - CSF Cut Dry & Packing")
                        lphtype = "LPHPrimaryCSFCutDryPacking";
                    else if (lphtype == "Intermediate Line - CSF Feeding & DCCC")
                        lphtype = "LPHPrimaryCSFInfeedConditioning";
                    else if (lphtype == "Intermediate Line - Clove Cut Dry & Packing")
                        lphtype = "LPHPrimaryCloveCutDryPacking";
                    else if (lphtype == "Intermediate Line - RTC")
                        lphtype = "LPHPrimaryRTC";
                    else if (lphtype == "Intermediate Line - Casing Kitchen")
                        lphtype = "LPHPrimaryKitchen";
                    else if (lphtype == "White Line OTP - Process Note")
                        lphtype = "LPHPrimaryWhiteLineOTP";
                    else if (lphtype == "Kretek Line - Feeding KR & RJ")
                        lphtype = "LPHPrimaryKretekLineFeeding";
                    else if (lphtype == "Kretek Line - DCCC KR & RJ")
                        lphtype = "LPHPrimaryKretekLineConditioning";
                    else if (lphtype == "Kretek Line - Cut Dry")
                        lphtype = "LPHPrimaryKretekLineCuttingDrying";
                    else if (lphtype == "Kretek Line - Packing")
                        lphtype = "LPHPrimaryKretekLinePacking";
                    else if (lphtype == "Kretek Line - CRES Feeding & DCCC")
                        lphtype = "LPHPrimaryCresFeedingConditioning";
                    else if (lphtype == "Kretek Line - CRES Cut Dry & Packing")
                        lphtype = "LPHPrimaryCresDryingPacking";
                    else if (lphtype == "White Line PMID - Feeding White")
                        lphtype = "LPHPrimaryWhiteLineFeedingWhite";
                    else if (lphtype == "White Line PMID - DCCC")
                        lphtype = "LPHPrimaryWhiteLineDCCC";
                    else if (lphtype == "White Line PMID - Cutting + FTD")
                        lphtype = "LPHPrimaryWhiteLineCuttingFTD";
                    else if (lphtype == "White Line PMID - Addback")
                        lphtype = "LPHPrimaryWhiteLineAddback";
                    else if (lphtype == "White Line PMID - Packing White")
                        lphtype = "LPHPrimaryWhiteLinePackingWhite";
                    else if (lphtype == "White Line PMID - Feeding SPM")
                        lphtype = "LPHPrimaryWhiteLineFeedingSPM";
                    else if (lphtype == "White Line PMID - Feeding IS White")
                        lphtype = "LPHPrimaryISWhiteFeeding";
                    else if (lphtype == "White Line PMID - Cut Dry IS White")
                        lphtype = "LPHPrimaryISWhiteCutDry";

                    lphSubs = lphSubs.Where(x => x.LPHHeader == lphtype + "Controller").ToList();
                }


                List<long> sudah = new List<long>();
                foreach (var item in approvalList.ToList())
                {
                    if (sudah.Contains(item.LPHSubmissionID))
                    {
                        approvalList.Remove(item); // chanif: hanya munculkan status approval terakhir
                        continue;
                    }

                    if (item.Status.Trim().ToLower() == "submitted" || item.Status.Trim() == "")
                    {
                        item.Status = "Waiting for Approval";
                        if (item.ApproverID == currentUserID || item.UserID == currentUserID)
                            item.IsNeedMyApproval = true;
                    }

                    if (locationMap.ContainsKey(item.LocationID))
                    {
                        string loc;
                        locationMap.TryGetValue(item.LocationID, out loc);
                        item.Location = loc;
                    }
                    else
                    {
                        item.Location = _locationAppService.GetLocationFullCode(item.LocationID);
                        locationMap.Add(item.LocationID, item.Location);
                    }

                    var lphsub = lphSubs.Where(x => x.ID == item.LPHSubmissionID).FirstOrDefault();
                    if (lphsub == null)
                    {
                        approvalList.Remove(item); // chanif: jika tidak ada datanya berarti draft, jangan munculkan approval
                        continue;
                    }
                    var check = lphs.Where(x => x.ID == lphsub.LPHID).FirstOrDefault();
                    // chanif: exclude LPH yang sudah dihapus
                    if (check == null)
                    {
                        approvalList.Remove(item);
                        continue;
                    }
                    else if (check.IsDeleted)
                    {
                        approvalList.Remove(item);
                        continue;
                    }

                    var temp = check.MenuTitle;
                    item.LPHType = temp.Replace("Controller", "");

                    //item.Approver = GetFullName(item.ApproverID);
                    item.User = GetCreatorBySubmissionID(item.LPHSubmissionID);
                    sudah.Add(item.LPHSubmissionID);
                }

                if (!string.IsNullOrEmpty(status))
                {
                    if (status.Trim().ToLower() == "approved")
                        approvalList = approvalList.Where(x => x.Status.Trim().ToLower() == "approved").ToList();
                }
                else
                {
                    approvalList = approvalList.Where(x => x.Status == "Waiting for Approval").ToList();
                }

                byte[] excelData = ExcelGenerator.ExportLPHPPApproval(approvalList, AccountName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=LPH-PP-Approval.xlsx");
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

        #region Helper Get Info
        private string GetFullName(long userID)
		{
			string user = _userAppService.GetById(userID);
			UserModel userModel = user.DeserializeToUser();

			string emp = _employeeAppService.GetBy("EmployeeID", userModel.EmployeeID, true);
			EmployeeModel empModel = emp.DeserializeToEmployee();

			return empModel.FullName;
		}

		private List<long> GetUserIDList(string employeeID)
		{
			List<long> result = new List<long>();
			string emp = _employeeAppService.GetAll();
			List<EmployeeModel> empList = emp.DeserializeToEmployeeList();

			// get level 1
			List<EmployeeModel> levelOneList = empList.Where(x => x.ReportToID1 != null && x.ReportToID1.Trim() == employeeID.Trim()).ToList();
			foreach (var item in levelOneList)
			{
				result.Add(item.ID);

				// get level 2
				List<EmployeeModel> levelTwoList = empList.Where(x => x.ReportToID1 != null && x.ReportToID1.Trim() == item.EmployeeID.Trim()).ToList();
				foreach (var lvl2 in levelTwoList)
				{
					result.Add(lvl2.ID);
				}
			}

			return result;
		}

		[HttpPost]
		public ActionResult GetPPLPHDetailBySUBSID(long subsID)
		{
			try
			{
				string subs = _ppLphSubmissionsAppService.GetById(subsID, true);
				PPLPHSubmissionsModel subsModel = subs.DeserializeToPPLPHSubmissions();

				string lph = _ppLphAppService.GetById(subsModel.LPHID);
				PPLPHModel model = lph.DeserializeToPPLPH();

				model.LPHType = model.MenuTitle.Replace("Controller", "");

				return Json(new { Status = "True", Result = model }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { Status = "False" }, JsonRequestBehavior.AllowGet);
			}
		}
        #endregion

        public List<SelectListItem> BindDropDownPPLPHType()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string lphs = _ppLphAppService.GetAll();
            List<PPLPHModel> lphList = lphs.DeserializeToPPLPHList();

            foreach (var item in lphList)
            {
                item.LPHType = item.MenuTitle.Replace("Controller", " ");
            }

            lphList = lphList.GroupBy(x => x.LPHType).Select(x => x.FirstOrDefault()).ToList();
            foreach (var item in lphList)
            {
                var text = item.LPHType.ToString().Trim();

                if (text == "LPHPrimaryKretekLineAddback")
                    text = "Kretek Line - Addback";
                else if (text == "LPHPrimaryDiet")
                    text = "Intermediate Line - DIET";
                else if (text == "LPHPrimaryCloveInfeedConditioning")
                    text = "Intermediate Line - Clove Feeding & DCCC";
                else if (text == "LPHPrimaryCSFCutDryPacking")
                    text = "Intermediate Line - CSF Cut Dry & Packing";
                else if (text == "LPHPrimaryCSFInfeedConditioning")
                    text = "Intermediate Line - CSF Feeding & DCCC";
                else if (text == "LPHPrimaryCloveCutDryPacking")
                    text = "Intermediate Line - Clove Cut Dry & Packing";
                else if (text == "LPHPrimaryRTC")
                    text = "Intermediate Line - RTC";
                else if (text == "LPHPrimaryKitchen")
                    text = "Intermediate Line - Casing Kitchen";
                else if (text == "LPHPrimaryWhiteLineOTP")
                    text = "White Line OTP - Process Note";
                else if (text == "LPHPrimaryKretekLineFeeding")
                    text = "Kretek Line - Feeding KR & RJ";
                else if (text == "LPHPrimaryKretekLineConditioning")
                    text = "Kretek Line - DCCC KR & RJ";
                else if (text == "LPHPrimaryKretekLineCuttingDrying")
                    text = "Kretek Line - Cut Dry";
                else if (text == "LPHPrimaryKretekLinePacking")
                    text = "Kretek Line - Packing";
                else if (text == "LPHPrimaryCresFeedingConditioning")
                    text = "Kretek Line - CRES Feeding & DCCC";
                else if (text == "LPHPrimaryCresDryingPacking")
                    text = "Kretek Line - CRES Cut Dry & Packing";
                else if (text == "LPHPrimaryWhiteLineFeedingWhite")
                    text = "White Line PMID - Feeding White";
                else if (text == "LPHPrimaryWhiteLineDCCC")
                    text = "White Line PMID - DCCC";
                else if (text == "LPHPrimaryWhiteLineCuttingFTD")
                    text = "White Line PMID - Cutting + FTD";
                else if (text == "LPHPrimaryWhiteLineAddback")
                    text = "White Line PMID - Addback";
                else if (text == "LPHPrimaryWhiteLinePackingWhite")
                    text = "White Line PMID - Packing White";
                else if (text == "LPHPrimaryWhiteLineFeedingSPM")
                    text = "White Line PMID - Feeding SPM";
                else if (text == "LPHPrimaryISWhiteFeeding")
                    text = "White Line PMID - Feeding IS White";
                else if (text == "LPHPrimaryISWhiteCutDry")
                    text = "White Line PMID - Cut Dry IS White";

                _menuList.Add(new SelectListItem
                {
                    Value = item.LPHType.ToString(),
                    Text = text
                });
            }
            return _menuList;
        }
        private LocationTreeModel GetLocationTreeModel()
        {
            LocationTreeModel model = new LocationTreeModel();

            int index = 1;

            // get production center list
            string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> pcList = pcs.DeserializeToLocationList();
            foreach (var item in pcList)
            {
                string pc = _referenceAppService.GetDetailBy("Code", item.Code, true);
                ProductionCenterModel pcModel = pc.DeserializeToProductionCenter(index++, item.ID, item.ParentID);
                model.ProductionCenters.Add(pcModel);
            }

            // get department list
            foreach (var pc in model.ProductionCenters)
            {
                LocationModel currentPC = pcList.Where(x => x.Code == pc.Code).FirstOrDefault();
                string departments = _locationAppService.FindBy("ParentID", currentPC.ID, true);
                List<LocationModel> departmentList = departments.DeserializeToLocationList();

                foreach (var d in departmentList)
                {
                    string depts = _referenceAppService.GetDetailBy("Code", d.Code, true);
                    DepartmentModel deptModel = depts.DeserializeToDepartment(index++, d.ID, d.ParentID);

                    string subdepartments = _locationAppService.FindBy("ParentID", d.ID, true);
                    List<LocationModel> subdepartmentList = subdepartments.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

                    foreach (var subdeb in subdepartmentList)
                    {
                        string sds = _referenceAppService.GetDetailBy("Code", subdeb.Code, true);
                        deptModel.SubDepartments.Add(sds.DeserializeToSubDepartment(index++, subdeb.ID, subdeb.ParentID));
                    }

                    pc.Departments.Add(deptModel);
                }
            }

            return model;
        }

        private string GetCreatorBySubmissionID(long submissionID)
        {
            string creator = "";
            string approval = _ppLphApprovalAppService.FindByNoTracking("LPHSubmissionID", submissionID.ToString(), true);
            List<PPLPHApprovalsModel> approveModel_list = approval.DeserializeToPPLPHApprovalList();
         
            var approveModel = approveModel_list.FirstOrDefault();
            creator = GetFullName(approveModel.UserID);

            return creator;

        }
    }
}