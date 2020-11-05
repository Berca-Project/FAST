using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
	public class CalendarHolidayService : ServiceBase<CalendarHoliday>, ICalendarHolidayService
	{
		public CalendarHolidayService(IUnitOfWork unitOfWork) : base(unitOfWork)
		{
		}
	}
}
