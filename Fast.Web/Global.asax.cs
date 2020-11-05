using System;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace Fast.Web
{
	public class MvcApplication : System.Web.HttpApplication
	{
		protected void Application_Start()
		{			
			AreaRegistration.RegisterAllAreas();
			FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
			RouteConfig.RegisterRoutes(RouteTable.Routes);
			BundleConfig.RegisterBundles(BundleTable.Bundles);
		}

        protected void Application_Error(Object sender, EventArgs e)
        {
            var exception = Server.GetLastError();
            if (exception != null)
            {
                Session["ErrorPage"] = exception.Message;
                Server.TransferRequest("~/Error");
            }
        }

		protected void Application_BeginRequest(object sender, EventArgs e)
		{
			//HttpCookie cookie = HttpContext.Current.Request.Cookies["Language"];
			//if (cookie != null && cookie.Value != null)
			//{
			//	System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(cookie.Value);
			//	System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(cookie.Value);
			//}
			//else
			//{
				System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en");
				System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en");
			//}
		}
	}
}
