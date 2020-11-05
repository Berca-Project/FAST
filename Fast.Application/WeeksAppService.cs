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
    public class WeeksAppService: AppServiceBase<Weeks>, IWeeksAppService
    {
        public WeeksAppService(IWeeksService serviceBase) : base(serviceBase)
        {
        }
    }
}
