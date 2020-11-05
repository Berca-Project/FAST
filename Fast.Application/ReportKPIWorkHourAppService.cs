using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Application
{
    public class ReportKPIWorkHourAppService: AppServiceBase<ReportKPIWorkHour>, IReportKPIWorkHourAppService
    {
        public ReportKPIWorkHourAppService(IReportKPIWorkHourService serviceBase) : base(serviceBase)
        {
        }
    }
}
