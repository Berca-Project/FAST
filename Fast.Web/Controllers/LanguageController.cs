using System.Globalization;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Controllers
{
	public class LanguageController : Controller
	{
		public ActionResult ChangeToID()
		{
			string language = "id";
			Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(language);
			Thread.CurrentThread.CurrentUICulture = new CultureInfo(language);

			HttpCookie cookie = new HttpCookie("Language");
			cookie.Value = language;

			Response.Cookies.Add(cookie);

			string previousUrl = System.Web.HttpContext.Current.Request.UrlReferrer.ToString();

			return Redirect(previousUrl);
		}

		public ActionResult ChangeToEN()
		{
			string language = "en";
			Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(language);
			Thread.CurrentThread.CurrentUICulture = new CultureInfo(language);

			HttpCookie cookie = new HttpCookie("Language");
			cookie.Value = language;

			Response.Cookies.Add(cookie);

			string previousUrl = System.Web.HttpContext.Current.Request.UrlReferrer.ToString();

			return Redirect(previousUrl);
		}
	}
}