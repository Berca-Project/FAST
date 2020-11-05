using Fast.Web.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Controllers.Master
{
	[AuthorizeAD]
	public class HomeController : Controller
	{
		public ActionResult Index()
		{
			if (Session["UserLogon"] == null)
			{
				return RedirectToAction("Index", "Login");
			}
			else
			{
				return View();
			}
        }

        public ActionResult Dashboard()
        {
            return View();
        }

        private class DashboardLocationDailyModel
        {
            public int ID { get; set; }
            public DateTime Date { get; set; }
            public int Submitted { get; set; }
            public int Approved { get; set; }
            public string Location { get; set; }
            public bool IsDeleted { get; set; }
            public DateTime ModifiedDate { get; set; }
        }

        private class DashboardLocationWeekModel
        {
            public int ID { get; set; }
            public int Week { get; set; }
            public int Year { get; set; }
            public int Submitted { get; set; }
            public int Approved { get; set; }
            public string Location { get; set; }
            public bool IsDeleted { get; set; }
            public DateTime ModifiedDate { get; set; }
        }

        private class DashboardLPHWeekViewModel
        {
            public int Week { get; set; }
            public int Year { get; set; }
            public int Approved { get; set; }
        }

        private class DashboardLPHDailyViewModel
        {
            public string Date { get; set; }
            public int Approved { get; set; }
        }

        private class DashboardLPHViewModel
        {
            public Dictionary<string, List<DashboardLPHDailyViewModel>> Dailies { get; set; }
            public Dictionary<string, List<DashboardLPHWeekViewModel>> Weeks { get; set; }
        }

        private class DashboardLocationModel
        {
            public string Location { get; set; }
            public double All { get; set; }
            public double Primary { get; set; }
            public double Secondary { get; set; }
            public double Trend { get; set; }
        }

        private class DashboardWasteModel
        {
            public int ID { get; set; }
            public double All { get; set; }
            public int Week { get; set; }
            public int Year { get; set; }
            public string Location { get; set; }
            public string Type { get; set; }
            public decimal Dry { get; set; }
            public decimal Wet { get; set; }
            public decimal Dust { get; set; }
            public decimal Hot { get; set; }
            public bool IsDeleted { get; set; }
            public DateTime ModifiedDate { get; set; }
        }
        private class DashboardWasteViewModel
        {
            public string Type { get; set; }
            public decimal Dry { get; set; }
            public decimal Wet { get; set; }
            public decimal Dust { get; set; }
            public decimal Hot { get; set; }
        }

        private class DashboardYieldGovernanceModel
        {
            public int ID { get; set; }
            public double All { get; set; }
            public int Week { get; set; }
            public int Year { get; set; }
            public string Location { get; set; }
            public string Area { get; set; }
            public string Type { get; set; }
            public decimal Target { get; set; }
            public decimal Value { get; set; }
            public bool IsDeleted { get; set; }
            public DateTime ModifiedDate { get; set; }
        }

        private class DashboardYieldGovernanceViewModel
        {
            public string Area { get; set; }
            public decimal DryYield { get; set; }
            public decimal WetYield { get; set; }
        }

        [HttpPost]
        public ActionResult LocationData(DateTime date)
        {
            try
            {
                //DateTime date = new DateTime(2020, 3, 2);
                DateTime yesterdayDate = date.AddDays(-1);

                int selectedYear = date.Year;
                int selectedWeek = GetWeek(date);
                int lastYear = selectedWeek == 1 ? selectedYear - 1 : selectedYear;
                int lastWeek = selectedWeek == 1 ? GetWeek(new DateTime(lastYear, 12, 31)) : selectedWeek - 1;

                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                string queryDailies = "SELECT * FROM [dbo].[Dashboards] WHERE [Date] BETWEEN '" + yesterdayDate.ToString("yyyy-MM-dd") + "' AND '" + date.ToString("yyyy-MM-dd") + "'";
                List<DashboardLocationDailyModel> dataDailies = GetData<List<DashboardLocationDailyModel>>(conString, queryDailies) ?? new List<DashboardLocationDailyModel>();
                string queryWeeks = "SELECT * FROM [dbo].[DashboardWeeks] WHERE ([Year] = '" + selectedYear + "' AND [Week] = '" + selectedWeek + "') OR ([Year] = '" + lastYear + "' AND [Week] = '" + lastWeek + "')";
                List<DashboardLocationWeekModel> dataWeeks = GetData<List<DashboardLocationWeekModel>>(conString, queryWeeks) ?? new List<DashboardLocationWeekModel>();

                List<DashboardLocationDailyModel> todayDailies = dataDailies.Where(x => x.Date == date).ToList();
                List<DashboardLocationDailyModel> yesterdayDailies = dataDailies.Where(x => x.Date == yesterdayDate).ToList();
                List<DashboardLocationWeekModel> thisWeeks = dataWeeks.Where(x => x.Year == selectedYear && x.Week == selectedWeek).ToList();
                List<DashboardLocationWeekModel> lastWeeks = dataWeeks.Where(x => x.Year == lastYear && x.Week == lastWeek).ToList();

                List<DashboardLocationModel> locationDailies = new List<DashboardLocationModel>();
                List<DashboardLocationModel> locationWeeks = new List<DashboardLocationModel>();
                foreach (string location in new List<string> { "All", "PI", "PJ", "PB", "PK" })
                {
                    List<DashboardLocationDailyModel> todayDailiesByLocation = todayDailies.Where(x => location == "All" || x.Location.Contains(location)).ToList();
                    List<DashboardLocationDailyModel> yesterdayDailiesByLocation = yesterdayDailies.Where(x => location == "All" || x.Location.Contains(location)).ToList();
                    List<DashboardLocationWeekModel> thisWeeksByLocation = thisWeeks.Where(x => location == "All" || x.Location.Contains(location)).ToList();
                    List<DashboardLocationWeekModel> lastWeeksByLocation = lastWeeks.Where(x => location == "All" || x.Location.Contains(location)).ToList();
                    Dictionary<string, Dictionary<string, double>> countDaily = new Dictionary<string, Dictionary<string, double>>() {
                        { "PP", new Dictionary<string, double>{ { "Submitted", 0 }, { "Approved", 0 } } },
                        { "SP", new Dictionary<string, double>{ { "Submitted", 0 }, { "Approved", 0 } } },
                        { "Last", new Dictionary<string, double>{ { "Submitted", 0 }, { "Approved", 0 } } }
                    };
                    Dictionary<string, Dictionary<string, double>> countWeekly = new Dictionary<string, Dictionary<string, double>>() {
                        { "PP", new Dictionary<string, double>{ { "Submitted", 0 }, { "Approved", 0 } } },
                        { "SP", new Dictionary<string, double>{ { "Submitted", 0 }, { "Approved", 0 } } },
                        { "Last", new Dictionary<string, double>{ { "Submitted", 0 }, { "Approved", 0 } } }
                    };
                    foreach (string department in new List<string> { "PP", "SP" })
                    {
                        List<DashboardLocationDailyModel> todayDailiesByDepartment = todayDailiesByLocation.Where(x => x.Location.Contains(department)).ToList();
                        List<DashboardLocationDailyModel> yesterdayDailiesByDepartment = yesterdayDailiesByLocation.Where(x => x.Location.Contains(department)).ToList();
                        countDaily[department]["Submitted"] += todayDailiesByDepartment.Sum(x => x.Submitted);
                        countDaily[department]["Approved"] += todayDailiesByDepartment.Sum(x => x.Approved);
                        countDaily["Last"]["Submitted"] += yesterdayDailiesByDepartment.Sum(x => x.Submitted);
                        countDaily["Last"]["Approved"] += yesterdayDailiesByDepartment.Sum(x => x.Approved);

                        List<DashboardLocationWeekModel> thisWeeksByDepartment = thisWeeksByLocation.Where(x => x.Location.Contains(department)).ToList();
                        List<DashboardLocationWeekModel> lastWeeksByDepartment = lastWeeksByLocation.Where(x => x.Location.Contains(department)).ToList();
                        countWeekly[department]["Submitted"] += thisWeeksByDepartment.Sum(x => x.Submitted);
                        countWeekly[department]["Approved"] += thisWeeksByDepartment.Sum(x => x.Approved);
                        countWeekly["Last"]["Submitted"] += lastWeeksByDepartment.Sum(x => x.Submitted);
                        countWeekly["Last"]["Approved"] += lastWeeksByDepartment.Sum(x => x.Approved);
                    }
                    double dailyAllSubmitted = countDaily["PP"]["Submitted"] + countDaily["SP"]["Submitted"];
                    double dailyAllApproved = countDaily["PP"]["Approved"] + countDaily["SP"]["Approved"];
                    double dailyAll = dailyAllSubmitted == 0 ? 0 : (dailyAllApproved / dailyAllSubmitted * 100);
                    double yesterdayAll = countDaily["Last"]["Submitted"] == 0 ? 0 : (countDaily["Last"]["Approved"] / countDaily["Last"]["Submitted"] * 100);
                    locationDailies.Add(new DashboardLocationModel
                    {
                        Location = location,
                        Primary = countDaily["PP"]["Submitted"] == 0 ? 0 : (countDaily["PP"]["Approved"] / countDaily["PP"]["Submitted"]) * 100,
                        Secondary = countDaily["SP"]["Submitted"] == 0 ? 0 : (countDaily["SP"]["Approved"] / countDaily["SP"]["Submitted"]) * 100,
                        All = dailyAll,
                        Trend = dailyAll - yesterdayAll,
                    });
                    double weeklyAllSubmitted = countWeekly["PP"]["Submitted"] + countWeekly["SP"]["Submitted"];
                    double weeklyAllApproved = countWeekly["PP"]["Approved"] + countWeekly["SP"]["Approved"];
                    double weeklyAll = weeklyAllSubmitted == 0 ? 0 : (weeklyAllApproved / weeklyAllSubmitted) * 100;
                    double lastWeekAll = countWeekly["Last"]["Submitted"] == 0 ? 0 : (countWeekly["Last"]["Approved"] / countWeekly["Last"]["Submitted"]) * 100;
                    locationWeeks.Add(new DashboardLocationModel
                    {
                        Location = location,
                        Primary = countWeekly["PP"]["Submitted"] == 0 ? 0 : (countWeekly["PP"]["Approved"] / countWeekly["PP"]["Submitted"]) * 100,
                        Secondary = countWeekly["SP"]["Submitted"] == 0 ? 0 : (countWeekly["SP"]["Approved"] / countWeekly["SP"]["Submitted"]) * 100,
                        All = weeklyAll,
                        Trend = weeklyAll - lastWeekAll,
                    });
                }

                return Json(new { Status = "True", locationDailies, locationWeeks }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult LPHData(DateTime date)
        {
            try
            {
                DateTime startOfWeek = DateToLast(date, DayOfWeek.Sunday);
                DateTime endOfWeek = DateToNext(date, DayOfWeek.Saturday);
                DateTime startOfLastWeek = CultureInfo.CurrentCulture.Calendar.AddWeeks(startOfWeek, -1);

                DateTime thisMonth = new DateTime(date.Year, date.Month, 1);
                DateTime lastMonth = CultureInfo.CurrentCulture.Calendar.AddWeeks(thisMonth, -4);
                string weeksCondition = "";
                for (int i = 0; i < 8; i++, lastMonth = CultureInfo.CurrentCulture.Calendar.AddWeeks(lastMonth, 1))
                {
                    if (i != 0) weeksCondition += " OR ";
                    weeksCondition += "([Year] = '" + lastMonth.Year + "' AND [Week] = '" + GetWeek(lastMonth) + "')";
                }
                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                string queryDailies = "SELECT * FROM [dbo].[Dashboards] WHERE [Date] BETWEEN '" + startOfLastWeek.ToString("yyyy-MM-dd") + "' AND '" + endOfWeek.ToString("yyyy-MM-dd") + "'";
                List<DashboardLocationDailyModel> dataDailies = GetData<List<DashboardLocationDailyModel>>(conString, queryDailies) ?? new List<DashboardLocationDailyModel>();
                string queryWeeks = "SELECT * FROM [dbo].[DashboardWeeks] WHERE " + weeksCondition;
                List<DashboardLocationWeekModel> dataWeeks = GetData<List<DashboardLocationWeekModel>>(conString, queryWeeks) ?? new List<DashboardLocationWeekModel>();
                Dictionary<string, DashboardLPHViewModel> lph = new Dictionary<string, DashboardLPHViewModel>();

                foreach (string location in new List<string> { "All", "PI", "PJ", "PB", "PK" })
                {
                    List<DashboardLocationDailyModel> dailiesByLocation = dataDailies.Where(x => location == "All" || x.Location.Contains(location)).ToList();
                    List<DashboardLocationWeekModel> weeksByLocation = dataWeeks.Where(x => location == "All" || x.Location.Contains(location)).ToList();
                    Dictionary<string, List<DashboardLPHDailyViewModel>> dailies = new Dictionary<string, List<DashboardLPHDailyViewModel>>() {
                        { "PP", new List<DashboardLPHDailyViewModel>() },
                        { "SP", new List<DashboardLPHDailyViewModel>() },
                    };
                    Dictionary<string, List<DashboardLPHWeekViewModel>> weeks = new Dictionary<string, List<DashboardLPHWeekViewModel>>() {
                        { "PP", new List<DashboardLPHWeekViewModel>() },
                        { "SP", new List<DashboardLPHWeekViewModel>() },
                    };

                    for (DateTime processedDate = startOfLastWeek; processedDate <= endOfWeek; processedDate = processedDate.AddDays(1))
                    {
                        List<DashboardLocationDailyModel> today = dailiesByLocation.Where(x => x.Date == processedDate).ToList();
                        foreach (string department in new List<string> { "PP", "SP" })
                        {
                            dailies[department].Add(new DashboardLPHDailyViewModel
                            {
                                Date = processedDate.ToString("yyyy-MM-dd"),
                                Approved = today.Where(x => x.Location.Contains(department)).Sum(x => x.Approved),
                            });
                        }
                    }

                    lastMonth = CultureInfo.CurrentCulture.Calendar.AddWeeks(thisMonth, -4);
                    for (int i = 0; i < 8; i++, lastMonth = CultureInfo.CurrentCulture.Calendar.AddWeeks(lastMonth, 1))
                    {
                        int thisWeek = GetWeek(lastMonth);
                        List<DashboardLocationWeekModel> thisWeeks = weeksByLocation.Where(x => x.Year == lastMonth.Year && x.Week == thisWeek).ToList();
                        foreach (string department in new List<string> { "PP", "SP" })
                        {
                            weeks[department].Add(new DashboardLPHWeekViewModel
                            {
                                Year = lastMonth.Year,
                                Week = thisWeek,
                                Approved = thisWeeks.Where(x => x.Location.Contains(department)).Sum(x => x.Approved),
                            });
                        }
                    }

                    lph.Add(location, new DashboardLPHViewModel
                    {
                        Dailies = dailies,
                        Weeks = weeks,
                    });
                }

                return Json(new { Status = "True", lph }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        private readonly static Dictionary<string, List<string>> wasteMap = new Dictionary<string, List<string>>
        {
            { "Intermediate", new List<string>
            {
                "CloveInfeedConditioning",
                "CloveCutDryPacking",
                "CSFInfeedConditioning",
                "CSFCutDryPacking",
                "RTC",
                "Diet",
            }},
            { "Kretek", new List<string>
            {
                "KretekLineFeeding",
                "KretekLineConditioning",
                "KretekLineCuttingDrying",
                "KretekLineAddback",
                "KretekLinePacking",
                "CresFeedingConditioning",
                "CresCutDryPacking",
                "Kitchen",
            }},
            { "White", new List<string>
            {
                "WhiteLineFeedingWhite",
                "WhiteLineDCCC",
                "WhiteLineCuttingFTD",
                "WhiteLineAddback",
                "WhiteLinePackingWhite",
                "WhiteLineFeedingSPM",
                "ISWhiteFeeding",
                "ISWhiteCutDry",
            }},
            { "WhiteOTP", new List<string>
            {
                "WhiteLineOTP",
            }},
        };

        [HttpPost]
        public ActionResult PrimaryReportData(DateTime date)
        {
            try
            {
                int week = GetWeek(date);

                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                string queryYG = "SELECT * FROM [dbo].[DashboardYieldGovernances] WHERE [Year] = " + date.Year + " AND [Week] = " + week + " AND [Type] IN ('WetYield', 'DryYield')";
                List<DashboardYieldGovernanceModel> dataYGs = GetData<List<DashboardYieldGovernanceModel>>(conString, queryYG) ?? new List<DashboardYieldGovernanceModel>();
                string queryWastes = "SELECT * FROM [dbo].[DashboardWastes] WHERE [Year] = " + date.Year + " AND [Week] = " + week;
                List<DashboardWasteModel> dataWastes = GetData<List<DashboardWasteModel>>(conString, queryWastes) ?? new List<DashboardWasteModel>();
                Dictionary<string, List<DashboardYieldGovernanceViewModel>> yieldGovernance = new Dictionary<string, List<DashboardYieldGovernanceViewModel>>();
                Dictionary<string, List<DashboardWasteViewModel>> waste = new Dictionary<string, List<DashboardWasteViewModel>>();

                foreach (KeyValuePair<string, string> location in new Dictionary<string, string> { { "All", "0" }, { "PJ", "3" }, { "PK", "4" }, { "PI", "5" }, { "PB", "6" } })
                {
                    List<DashboardYieldGovernanceModel> ygByLocation = dataYGs.Where(x => location.Key == "All" || x.Location == location.Value).ToList();
                    List<DashboardWasteModel> wasteByLocation = dataWastes.Where(x => location.Key == "All" || x.Location == location.Value).ToList();
                    List<DashboardYieldGovernanceViewModel> ygs = new List<DashboardYieldGovernanceViewModel>();
                    List<DashboardWasteViewModel> wastes = new List<DashboardWasteViewModel>();

                    foreach (string area in new List<string> { "Clove", "White", "Kretek", "Diet" })
                    {
                        List<DashboardYieldGovernanceModel> ygThisArea = ygByLocation.Where(x => x.Area == area).ToList();
                        ygs.Add(new DashboardYieldGovernanceViewModel
                        {
                            Area = area,
                            WetYield = ygThisArea.Where(x => x.Type == "WetYield").FirstOrDefault()?.Value ?? 0,
                            DryYield = ygThisArea.Where(x => x.Type == "DryYield").FirstOrDefault()?.Value ?? 0,
                        });
                    }
                    foreach (string area in new List<string> { "Intermediate", "Kretek", "White", "WhiteOTP" })
                    {
                        List<string> map = wasteMap[area];
                        List<DashboardWasteModel> wasteThisArea = wasteByLocation.Where(x => map.Any(y => y == x.Type)).ToList();
                        wastes.Add(new DashboardWasteViewModel
                        {
                            Type = area,
                            Dry = wasteThisArea.Sum(x => x.Dry),
                            Wet = wasteThisArea.Sum(x => x.Wet),
                            Dust = wasteThisArea.Sum(x => x.Dust),
                            Hot = wasteThisArea.Sum(x => x.Hot),
                        });
                    }
                    yieldGovernance.Add(location.Key, ygs);
                    waste.Add(location.Key, wastes);
                }

                return Json(new { Status = "True", yieldGovernance, waste }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        private class DashboardLPHNotificationModel
        {
            public DateTime Date { get; set; }
            public string Location { get; set; }
            public int Unapproved { get; set; }
        }

        [HttpPost]
        public ActionResult LPHNotificationData(DateTime date)
        {
            try
            {
                DateTime startOfWeek = DateToLast(date, DayOfWeek.Sunday);
                string sof = startOfWeek.ToString("yyyy-MM-dd");
                DateTime endOfWeek = DateToNext(date, DayOfWeek.Saturday);
                string eof = endOfWeek.ToString("yyyy-MM-dd");

                string conString = ConfigurationManager.ConnectionStrings["FastAppConn"].ConnectionString;
                string querySp = @"SELECT [LPHSubmissions].[Date], [LPHSubmissions].[Location], COUNT(*) AS [Unapproved]
                    FROM [dbo].[LPHApprovals] AS [Approval]
                    RIGHT JOIN (
	                    SELECT MAX([ID]) AS [ID]
	                    FROM [dbo].[LPHApprovals]
	                    WHERE 
		                    [Date] BETWEEN '" + sof + "' AND '" + eof + @"'
                        GROUP BY [LPHSubmissionID]
                    ) AS [IDs] ON [Approval].[ID] = [IDs].[ID]
                    LEFT JOIN [dbo].[LPHSubmissions] ON [LPHSubmissions].[ID] = [LPHSubmissionID]
                    WHERE 
	                    [Status] = 'Submitted' 
	                    AND [LPHSubmissions].[Date] IS NOT NULL 
	                    AND [LPHSubmissions].[Location] IS NOT NULL
	                    AND [LPHSubmissions].[Location] != '-'
                    GROUP BY [LPHSubmissions].[Date], [LPHSubmissions].[Location]
                    ORDER BY [Date]";
                List<DashboardLPHNotificationModel> dataSps = GetData<List<DashboardLPHNotificationModel>>(conString, querySp) ?? new List<DashboardLPHNotificationModel>();
                string queryPp = @"SELECT [PPLPHSubmissions].[Date], [PPLPHSubmissions].[Location], COUNT(*) AS [Unapproved]
                    FROM [dbo].[PPLPHApprovals] AS [Approval]
                    RIGHT JOIN (
	                    SELECT MAX([ID]) AS [ID]
	                    FROM [dbo].[PPLPHApprovals]
	                    WHERE 
		                    [Date] BETWEEN '" + sof + "' AND '" + eof + @"'
	                    GROUP BY [LPHSubmissionID]
                    ) AS [IDs] ON [Approval].[ID] = [IDs].[ID]
                    LEFT JOIN [dbo].[PPLPHSubmissions] ON [PPLPHSubmissions].[ID] = [LPHSubmissionID]
                    WHERE 
	                    [Status] = 'Submitted' 
	                    AND [PPLPHSubmissions].[Date] IS NOT NULL 
	                    AND [PPLPHSubmissions].[Location] IS NOT NULL
	                    AND [PPLPHSubmissions].[Location] != '-'
                    GROUP BY [PPLPHSubmissions].[Date], [PPLPHSubmissions].[Location]
                    ORDER BY [Date]";
                List<DashboardLPHNotificationModel> dataPps = GetData<List<DashboardLPHNotificationModel>>(conString, queryPp) ?? new List<DashboardLPHNotificationModel>();
                Dictionary<string, Dictionary<string, int>> lphSp = new Dictionary<string, Dictionary<string, int>>();
                Dictionary<string, Dictionary<string, int>> lphPp = new Dictionary<string, Dictionary<string, int>>();

                foreach (string location in new List<string> { "All", "PI", "PJ", "PB", "PK" })
                {
                    List<DashboardLPHNotificationModel> spByLocation = dataSps.Where(x => location == "All" || x.Location == location).ToList();
                    List<DashboardLPHNotificationModel> ppByLocation = dataPps.Where(x => location == "All" || x.Location == location).ToList();
                    Dictionary<string, int> sps = new Dictionary<string, int>();
                    Dictionary<string, int> pps = new Dictionary<string, int>();

                    for (DateTime processedDate = startOfWeek; processedDate <= endOfWeek; processedDate = processedDate.AddDays(1))
                    {
                        string d = processedDate.ToString("yyyy-MM-dd");
                        DashboardLPHNotificationModel sp = spByLocation.Where(x => x.Date == processedDate).FirstOrDefault();
                        DashboardLPHNotificationModel pp = ppByLocation.Where(x => x.Date == processedDate).FirstOrDefault();

                        sps.Add(d, sp?.Unapproved ?? 0);
                        pps.Add(d, pp?.Unapproved ?? 0);
                    }

                    lphSp.Add(location, sps);
                    lphPp.Add(location, pps);
                }

                return Json(new { Status = "True", lphSp, lphPp }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Status = "False", Error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        private static DateTime DateToLast(DateTime date, DayOfWeek dayOfWeek) => date.AddDays(-((7 + (date.DayOfWeek - dayOfWeek)) % 7));
        private static DateTime DateToNext(DateTime date, DayOfWeek dayOfWeek) => date.AddDays(((7 + (dayOfWeek - date.DayOfWeek)) % 7));
        private static int GetWeek(DateTime d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);

		private static T GetData<T>(string conString, string query)
		{
			DataSet dset = new DataSet();
			using (SqlConnection con = new SqlConnection(conString))
			{
				SqlCommand cmd = new SqlCommand(query, con);
				using (SqlDataAdapter da = new SqlDataAdapter(cmd))
				{
					da.Fill(dset);
				}
			}

			return JsonConvert.SerializeObject(dset.Tables[0]).DeserializeJson<T>();
		}
    }
}