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
    public class ReportKPIDIMAppService: AppServiceBase<ReportKPIDIM>, IReportKPIDIMAppService
    {
        public ReportKPIDIMAppService(IReportKPIDIMService serviceBase) : base(serviceBase)
        {
        }
    }
}
