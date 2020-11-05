using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class CalendarHolidayAppService : AppServiceBase<CalendarHoliday>, ICalendarHolidayAppService
	{
		public CalendarHolidayAppService(ICalendarHolidayService serviceBase) : base(serviceBase)
		{
		}
	}
}