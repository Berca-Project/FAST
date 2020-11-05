using System;
namespace Fast.Infra.CrossCutting.Common
{
	public static class DateTimeExtensions
	{
		public static DateTime StartOfWeek(this DateTime dt)
		{
			int diff = (7 + (dt.DayOfWeek - DayOfWeek.Monday)) % 7;
			return dt.AddDays(-1 * diff).Date;
		}

		public static DateTime EndOfWeek(this DateTime dt)
		{
			int diff = (7 + (dt.DayOfWeek - DayOfWeek.Monday)) % 7;
			DateTime startDate = dt.AddDays(-1 * diff).Date;

			return startDate.AddDays(6);
		}

		public static DateTime StartDateByWeekNumber(this DateTime dateToCheck, int week)
		{
			DateTime dt = DateTime.Now;
			int diff = (7 + (dt.DayOfWeek - DayOfWeek.Monday)) % 7;
			DateTime startDate = dt.AddDays(-1 * diff).Date;
			week -= 1;

			return startDate.AddDays(week * 7);			
		}

		public static DateTime EndDateByWeekNumber(this DateTime dateToCheck, int week)
		{
			DateTime dt = DateTime.Now;
			int diff = (7 + (dt.DayOfWeek - DayOfWeek.Monday)) % 7;
			DateTime startDate = dt.AddDays(-1 * diff).Date;
			week -= 1;
			
			return startDate.AddDays(week * 13);
		}

		public static bool IsInTheTargetWeekRange(this DateTime dateToCheck, int week)
		{
			DateTime dt = DateTime.Now;
			int diff = (7 + (dt.DayOfWeek - DayOfWeek.Monday)) % 7;
			DateTime startDate = dt.AddDays(-1 * diff).Date;
			week -= 1;

			DateTime startDateNextWeek = startDate.AddDays(week * 7);
			DateTime endDateNextWeek = startDate.AddDays(week * 13);

			return dateToCheck >= startDateNextWeek && dateToCheck <= endDateNextWeek;
		}
	}
}
