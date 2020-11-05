using Fast.Web.Utils;
using System.Web.Mvc;

namespace Fast.Web.Controllers
{
	public class HelpController : Controller
	{
		public ActionResult UserGuide()
		{
			if (Session["UserLogon"] == null)
			{
				return RedirectToAction("Index", "Login");
			}
			return View();
		}

		public ActionResult TrainingMaterial()
		{
			if (Session["UserLogon"] == null)
			{
				return RedirectToAction("Index", "Login");
			}
			return View();
		}

		public ActionResult Infographic()
		{
			if (Session["UserLogon"] == null)
			{
				return RedirectToAction("Index", "Login");
			}
			return View();
		}

		public ActionResult MandatoryBible()
		{
			if (Session["UserLogon"] == null)
			{
				return RedirectToAction("Index", "Login");
			}
			return View();
		}

		public ActionResult PIC()
		{
			if (Session["UserLogon"] == null)
			{
				return RedirectToAction("Index", "Login");
			}
			return View();
		}
	}
}