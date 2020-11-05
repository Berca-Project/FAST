using System.Web.Mvc;

namespace Fast.Web.Controllers
{
    public class ErrorController : Controller
    {
        public ActionResult Index()
        {            
            ViewBag.Message = Session["ErrorPage"];
            
            return View();
        }
    }
}