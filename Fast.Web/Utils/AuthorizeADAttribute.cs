using System.Configuration;
using System.Web;
using System.Web.Mvc;

namespace Fast.Web.Utils
{
	public class AuthorizeADAttribute : AuthorizeAttribute
	{
		private bool _authenticated;
		private bool _authorized;

		public string Groups
		{
			get { return ConfigurationManager.AppSettings["GroupAD"].ToString(); }
		}

		protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
		{
			base.HandleUnauthorizedRequest(filterContext);

			if (!_authenticated || !_authorized)
			{
				var baseUrl = ConfigurationManager.AppSettings["BaseUrl"].ToString();
				bool byPass = bool.Parse(ConfigurationManager.AppSettings["ByPass"] ?? "false");

				//bypass = true tanpa AD
				if (byPass)
				{
					filterContext.Result = new RedirectResult(baseUrl + "/Login");
				}
				else
				{
					filterContext.Result = new RedirectResult(baseUrl + "/Unauthorized");
				}
			}
		}

		protected override bool AuthorizeCore(HttpContextBase httpContext)
		{
			bool usePageAuth = bool.Parse(ConfigurationManager.AppSettings["UsePageAuth"] ?? "false");

			string userku = httpContext.User.Identity.Name;
			string myDomain = ConfigurationManager.AppSettings["MyDomain"];
			userku = userku.Replace(myDomain + "\\", "");
			userku = myDomain + "\\" + userku;

			if (!usePageAuth || userku.ToLower().Contains("\\f-"))
				return true;

			_authenticated = base.AuthorizeCore(httpContext);

			if (_authenticated)
			{
				if (string.IsNullOrEmpty(Groups))
				{
					_authorized = true;
					return _authorized;
				}

				var groups = Groups.Split(',');

				try
				{
					foreach (var group in groups)
					{
						if (httpContext.User.IsInRole(group))
						{
							return true;
						}
					}
				}
				catch
				{
					_authorized = false;
					return _authorized;
				}
			}

			_authorized = false;
			return _authorized;
		}
	}
}