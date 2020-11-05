using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Utils;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
	[CustomAuthorize("employeeleave")]
	public class EmployeeLeaveController : BaseController<EmployeeLeaveModel>
	{

        private readonly IReferenceAppService _referenceAppService;
        private readonly IEmployeeAppService _empService;
		private readonly IEmployeeLeaveAppService _empLeaveService;
		private readonly IMenuAppService _menuService;
		private readonly IUserAppService _userAppService;
		private readonly ILoggerAppService _logger;
        private readonly ILocationAppService _locationAppService;

        public EmployeeLeaveController(
			IEmployeeAppService empService,
			ILoggerAppService logger,
			IUserAppService userAppService,
			IEmployeeLeaveAppService empLeaveService,
            ILocationAppService locationAppService,
            IReferenceAppService referenceAppService,
            IMenuAppService menuService)
		{
            _referenceAppService = referenceAppService;
            _locationAppService = locationAppService;
            _empService = empService;
			_empLeaveService = empLeaveService;
			_menuService = menuService;
			_userAppService = userAppService;
			_logger = logger;
		}

		// GET: EmployeeLeave
		public ActionResult Index()
		{
			EmployeeLeaveModel model = new EmployeeLeaveModel();
			model.Access = GetAccess(WebConstants.MenuSlug.EMP_LEAVE, _menuService);
            ViewBag.ProductionCenterList = DropDownHelper.GetProductionCenterInIndonesia(_locationAppService, _referenceAppService);

            return View(model);
		}

		public ActionResult GenerateExcel(string startDateParam, string endDateParam, long locID, string locType)
		{
			try
			{
				// Getting all data    			
				string empLeaveList = _empLeaveService.GetAll();
				List<EmployeeLeaveModel> empLeaveModelList = empLeaveList.DeserializeToEmployeeLeaveList().Where(x => x.EmployeeID != null).ToList();

				string empList = _empService.GetAll(true);
				List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

				if (!string.IsNullOrEmpty(startDateParam))
				{
					DateTime startDate = DateTime.Parse(startDateParam);
					empLeaveModelList = empLeaveModelList.Where(x => x.StartDate.HasValue && x.StartDate.Value >= startDate && !string.IsNullOrEmpty(x.Location)).ToList();
				}

				if (!string.IsNullOrEmpty(endDateParam))
				{
					DateTime endDate = DateTime.Parse(endDateParam);
					empLeaveModelList = empLeaveModelList.Where(x => x.EndDate.HasValue && x.EndDate.Value <= endDate && !string.IsNullOrEmpty(x.Location)).ToList();
				}

				string userList = _userAppService.FindBy("IsFast", "true");
				List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();

				List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();


                empLeaveModelList = empLeaveModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).ToList();

                // Filter Search
                if(locID>0 && locType != null)
                {
                    List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);
                    empLeaveModelList = empLeaveModelList.Where(x => !string.IsNullOrEmpty(x.Location) && locationIdList.Any(y => y == x.LocationID)).OrderBy(x => x.StartDate).ToList();
                }

                foreach (var item in empLeaveModelList)
				{
					EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
					item.FullName = emp == null ? string.Empty : emp.FullName;
				}

				byte[] excelData = ExcelGenerator.ExportEmployeeLeave(empLeaveModelList, AccountName);

				Response.Clear();
				Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
				Response.AddHeader("content-disposition", "attachment;filename=Planning-Employee-Leaves.xlsx");
				Response.BinaryWrite(excelData);
				Response.End();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);
			}

			return RedirectToAction("Index");
		}

		public ActionResult GeneratePDF(string startDateParam, string endDateParam)
		{
			try
			{
				// Getting all data    			
				string empLeaveList = _empLeaveService.GetAll();
				List<EmployeeLeaveModel> empLeaveModelList = empLeaveList.DeserializeToEmployeeLeaveList().Where(x => x.EmployeeID != null).ToList();

				string empList = _empService.GetAll(true);
				List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

				if (!string.IsNullOrEmpty(startDateParam))
				{
					DateTime startDate = DateTime.Parse(startDateParam);
					empLeaveModelList = empLeaveModelList.Where(x => x.StartDate.HasValue && x.StartDate.Value >= startDate).ToList();
				}

				if (!string.IsNullOrEmpty(endDateParam))
				{
					DateTime endDate = DateTime.Parse(endDateParam);
					empLeaveModelList = empLeaveModelList.Where(x => x.EndDate.HasValue && x.EndDate.Value <= endDate).ToList();
				}

				string userList = _userAppService.FindBy("IsFast", "true");
				List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();

				List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();

				empLeaveModelList = empLeaveModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).OrderBy(x => x.StartDate).ToList();

				foreach (var item in empLeaveModelList)
				{
					EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
					item.FullName = emp == null ? string.Empty : emp.FullName;
				}

				Document pdfDoc = new Document(PageSize.A3.Rotate(), 10, 10, 10, 10);
				PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
				pdfDoc.Open();

				Image image = Image.GetInstance(Server.MapPath("~/Content/theme/images/fast-blue.jpg"));
				image.ScaleAbsolute(193, 38);
				pdfDoc.Add(image);

				BaseFont bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.EMBEDDED);
				Font font = new Font(bf, 12);
				pdfDoc.Add(new Paragraph(new Chunk("Title          : Master Employee Leaves", font)));
				pdfDoc.Add(new Paragraph(new Chunk("Generated By   : " + AccountName, font)));
				pdfDoc.Add(new Paragraph(new Chunk("Generated Date : " + DateTime.Now.ToString("dd-MMM-yy HH:mm:ss"), font)));

				//Horizontal Line
				Paragraph line = new Paragraph(new Chunk(new LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 1)));
				pdfDoc.Add(line);

				//Table
				PdfPTable table = new PdfPTable(11);
				table.WidthPercentage = 100;
				//0=Left, 1=Centre, 2=Right
				table.HorizontalAlignment = 0;
				table.SpacingBefore = 20f;
				table.SpacingAfter = 30f;

				PdfPCell cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				Chunk chunk = new Chunk("Emp ID");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("Emp Name");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("Emp Type");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("Leave Type");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("Start Date");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("End Date");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("Start HalfDay");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("Start Shift");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("End HalfDay");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("End Shift");
				cell.AddElement(chunk);
				table.AddCell(cell);

				cell = new PdfPCell();
				cell.BackgroundColor = Color.GRAY;
				chunk = new Chunk("Status");
				cell.AddElement(chunk);
				table.AddCell(cell);

				foreach (var item in empLeaveModelList)
				{
					table.AddCell(item.EmployeeID);
					table.AddCell(item.FullName);
					table.AddCell(item.EmployeeType);
					table.AddCell(item.LeaveType);
					table.AddCell(item.StartDate.HasValue ? item.StartDate.Value.ToString("dd-MMM-yy") : string.Empty);
					table.AddCell(item.EndDate.HasValue ? item.EndDate.Value.ToString("dd-MMM-yy") : string.Empty);
					table.AddCell(item.StartDateHalfDay);
					table.AddCell(item.StartDatePagiSiang);
					table.AddCell(item.EndDateHalfDay);
					table.AddCell(item.EndDatePagiSiang);
					table.AddCell(item.Status);
				}

				pdfDoc.Add(table);
				pdfWriter.CloseStream = false;
				pdfDoc.Close();

				Response.Buffer = true;
				Response.ContentType = "application/pdf";
				Response.AddHeader("content-disposition", "attachment;filename=Master-Employee-Leaves.pdf");
				Response.Cache.SetCacheability(HttpCacheability.NoCache);
				Response.Write(pdfDoc);
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
				string empLeaveList = _empLeaveService.GetAll();
				List<EmployeeLeaveModel> empLeaveModelList = empLeaveList.DeserializeToEmployeeLeaveList();
				empLeaveModelList = empLeaveModelList.OrderBy(x => x.StartDate).ToList();

				string empList = _empService.GetAll(true);
				List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

				string userList = _userAppService.FindBy("IsFast", "true");
				List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();

				List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();

				empLeaveModelList = empLeaveModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).OrderBy(x => x.StartDate).ToList();

				int recordsTotal = empLeaveModelList.Count();
				sortColumn = sortColumn == "ID" ? "" : sortColumn;
				bool isLoaded = false;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<string, string> empMap = new Dictionary<string, string>();
					foreach (var item in empLeaveModelList)
					{
                        if (empMap.ContainsKey(item.EmployeeID.Trim()))
                        {
                            string name;
                            empMap.TryGetValue(item.EmployeeID.Trim(), out name);
                            item.FullName = name;
                        }
                        else
                        {
                            EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
                            item.FullName = emp == null ? string.Empty : emp.FullName;
                            empMap.Add(item.EmployeeID.Trim(), item.FullName);
                        }
					}
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					empLeaveModelList = empLeaveModelList.Where(m => m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
																m.FullName.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "employeeid":
								empLeaveModelList = empLeaveModelList.OrderBy(x => x.EmployeeID).ToList();
								break;
							case "fullname":
								empLeaveModelList = empLeaveModelList.OrderBy(x => x.FullName).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "employeeid":
								empLeaveModelList = empLeaveModelList.OrderByDescending(x => x.EmployeeID).ToList();
								break;
							case "fullname":
								empLeaveModelList = empLeaveModelList.OrderByDescending(x => x.FullName).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = empLeaveModelList.Count();

				// Paging     
				var data = empLeaveModelList.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					foreach (var item in data)
					{
						EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
						item.FullName = emp == null ? string.Empty : emp.FullName;
					}
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<EmployeeLeaveModel>() }, JsonRequestBehavior.AllowGet);
			}
		}

		[HttpPost]
		public ActionResult GetAllByParam(string startDateParam, string endDateParam, long locID, string locType)
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
                string empLeaveList = _empLeaveService.GetAll();
                List<EmployeeLeaveModel> empLeaveModelList = empLeaveList.DeserializeToEmployeeLeaveList().Where(x => x.EmployeeID != null).ToList();

                string empList = _empService.GetAll(true);
				List<EmployeeModel> empModelList = empList.DeserializeToEmployeeList();

				if (!string.IsNullOrEmpty(startDateParam))
				{
					DateTime startDate = DateTime.Parse(startDateParam);
					empLeaveModelList = empLeaveModelList.Where(x => x.StartDate.HasValue && x.StartDate.Value >= startDate).ToList();
				}

				if (!string.IsNullOrEmpty(endDateParam))
				{
					DateTime endDate = DateTime.Parse(endDateParam);
					empLeaveModelList = empLeaveModelList.Where(x => x.EndDate.HasValue && x.EndDate.Value <= endDate).ToList();
				}

				string userList = _userAppService.FindBy("IsFast", "true");
				List<UserModel> userModelList = userList.DeserializeToUserList().Where(x => x.EmployeeID != null).ToList();

				List<string> empIDFastList = userModelList.Select(x => x.EmployeeID).ToList();
                if (locID > 0 && locType != null)
                {
                    List<long> locationIdList = _locationAppService.GetLocIDListByLocType(locID, locType);
                    empLeaveModelList = empLeaveModelList.Where(x => !string.IsNullOrEmpty(x.Location) && locationIdList.Any(y => y == x.LocationID)).ToList();
                }

                empLeaveModelList = empLeaveModelList.Where(x => empIDFastList.Any(y => y.Trim() == x.EmployeeID.Trim())).OrderBy(x => x.StartDate).ToList();

				int recordsTotal = empLeaveModelList.Count();

				sortColumn = sortColumn == "ID" ? "" : sortColumn;
				bool isLoaded = false;

				if (!string.IsNullOrEmpty(searchValue) || !string.IsNullOrEmpty(sortColumn))
				{
					isLoaded = true;
					Dictionary<long, string> locationMap = new Dictionary<long, string>();
					foreach (var item in empLeaveModelList)
					{
						EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
						item.FullName = emp == null ? string.Empty : emp.FullName;
					}
				}

				// Search    
				if (!string.IsNullOrEmpty(searchValue))
				{
					empLeaveModelList = empLeaveModelList.Where(m => m.EmployeeID.ToLower().Contains(searchValue.ToLower()) ||
																m.FullName.ToLower().Contains(searchValue.ToLower())).ToList();
				}

				if (!string.IsNullOrEmpty(sortColumn))
				{
					if (sortColumnDir == "asc")
					{
						switch (sortColumn.ToLower())
						{
							case "employeeid":
								empLeaveModelList = empLeaveModelList.OrderBy(x => x.EmployeeID).ToList();
								break;
							case "fullname":
								empLeaveModelList = empLeaveModelList.OrderBy(x => x.FullName).ToList();
								break;
							default:
								break;
						}
					}
					else
					{
						switch (sortColumn.ToLower())
						{
							case "employeeid":
								empLeaveModelList = empLeaveModelList.OrderByDescending(x => x.EmployeeID).ToList();
								break;
							case "fullname":
								empLeaveModelList = empLeaveModelList.OrderByDescending(x => x.FullName).ToList();
								break;
							default:
								break;
						}
					}
				}

				// total number of rows count     
				int recordsFiltered = empLeaveModelList.Count();

				// Paging     
				var data = empLeaveModelList.Skip(skip).Take(pageSize).ToList();

				if (!isLoaded)
				{
					foreach (var item in data)
					{
						EmployeeModel emp = empModelList.Where(x => x.EmployeeID.Trim().Equals(item.EmployeeID.Trim())).FirstOrDefault();
						item.FullName = emp == null ? string.Empty : emp.FullName;
					}
				}

				// Returning Json Data    
				return Json(new { data, recordsFiltered, recordsTotal, draw }, JsonRequestBehavior.AllowGet);
			}
			catch (Exception ex)
			{
				ViewBag.Result = false;

				_logger.LogError(ex.GetAllMessages(), AccountID, AccountName);

				return Json(new { data = new List<EmployeeLeaveModel>() }, JsonRequestBehavior.AllowGet);
			}
		}
	}
}
