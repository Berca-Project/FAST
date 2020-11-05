using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;

namespace Fast.Domain.Services
{
    public class ReportKPIWorkHourService: ServiceBase<ReportKPIWorkHour>, IReportKPIWorkHourService
    {
        public ReportKPIWorkHourService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
