using Fast.Web.Models;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.Web.Routing;

namespace Fast.Web.Utils
{
	/// <summary>
	/// Custom authorize attribute class
	/// </summary>
	public class CustomAuthorizeAttribute : AuthorizeAttribute
	{
		private string[] _menuSlugs;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="privilege_name"></param>
		public CustomAuthorizeAttribute(params string[] menuSlugs)
		{
			this._menuSlugs = menuSlugs;
		}

		/// <summary>
		/// Override the onAuthorization
		/// </summary>
		/// <param name="filterContext"></param>
		public override void OnAuthorization(AuthorizationContext filterContext)
		{
			var account = (UserModel)filterContext.HttpContext.Session["UserLogon"];
			if (account == null)
			{
				// redirect to login page if user not logged in  
				filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary { { "action", "Index" }, { "controller", "Login" } });
			}
			else
			{
				bool IsAuthorized = false;
				List<MenuModel> userAuthMenus = (List<MenuModel>)filterContext.HttpContext.Session["AuthMenu"];

				foreach (var menuslug in _menuSlugs)
				{
					IsAuthorized = userAuthMenus.Any(x => x.PageSlug.Equals(menuslug));
					if (!IsAuthorized)
						break;
				}

				if (!IsAuthorized)
				{
					this.HandleUnauthorizedRequest(filterContext);
				}
			}
		}

		/// <summary>
		/// Override the Handle of unauthorized request
		/// </summary>
		/// <param name="filterContext"></param>
		protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
		{
			filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary { { "action", "Index" }, { "controller", "Unauthorized" } });
		}
	}
}