using System.Web.Mvc;

namespace Fast.Web.Controllers
{
	public class UnauthorizedController : Controller
	{
		// GET: Unauthorized
		public ActionResult Index()
		{
			return View();
		}
	}
}