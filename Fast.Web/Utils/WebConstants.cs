using System.Configuration;

namespace Fast.Web.Utils
{
	public static class WebConstants
	{
		public static string BASEURLKU = ConfigurationManager.AppSettings["BaseUrl"];
        public static string DEFAULT_STICK = "Mio Stick";

        public static class MenuSlug
		{
			public static string USER = "user";
			public static string JOB_TITLE = "jobtitle";
			public static string ROLE = "role";
			public static string ACCESS_RIGHT = "accessright";
			public static string EMP_OVERTIME = "employeeovertime";
			public static string EMP_LEAVE = "employeeleave";
			public static string EMPLOYEE = "employee";
			public static string EMPLOYEE_ALL = "peoplesoft";
			public static string EMP_SKILL = "employeeskill";
			public static string LOCATION = "location";
			public static string MACHINE = "machine";
			public static string REFERENCE = "reference";
			public static string WPP = "wpp";
			public static string CALENDAR_UPLOAD = "calendarupload";
			public static string CALENDAR_HOLIDAY = "calendarholiday";
			public static string CALENDAR_REPORT = "calendarreport";
			public static string USER_MACHINE = "usermachines";
			public static string USER_LOGS = "userlogs";
			public static string MENU = "menu";
			public static string MPP = "mpp";
			public static string TRAINING = "training";
			public static string BRAND = "brand";
			public static string BLEND = "blend";
			public static string USER_ROLE = "userroles";
			public static string MATERIAL_CODE = "materialcode";
			public static string BRAND_CONVERSION = "brandconversion";
			public static string MAN_POWER = "manpower";
			public static string MEAL_REQUEST = "mealrequest";
			public static string SHUTTLE_REQUEST = "shuttlerequest";            
        }
	}
}

