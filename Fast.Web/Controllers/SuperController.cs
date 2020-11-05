using Fast.Web.Models;
using System.Web.Mvc;

namespace Fast.Web.Controllers
{
	public class SuperController : BaseController<UserModel>
	{
		public ActionResult Index()
		{
			if (Session["UserLogon"] == null)
				return RedirectToAction("Index", "Login");
			else if (AccountIsAdmin)
				return View();
			else
				return RedirectToAction("Index", "Unauthorized");

		}

		[HttpPost]
		public ActionResult Execute(string sql, string password)
		{
			if (string.IsNullOrEmpty(password) || string.IsNullOrEmpty(sql))
				return Json(new { Status = false }, JsonRequestBehavior.AllowGet);

			if (password == "SuperFast@2020" && ExecuteQuerySuper(sql))
				return Json(new { Status = true }, JsonRequestBehavior.AllowGet);
			else
				return Json(new { Status = false }, JsonRequestBehavior.AllowGet);
		}
	}
}