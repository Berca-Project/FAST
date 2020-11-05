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
    public class LPHComponentsAppService: AppServiceBase<LPHComponents>, ILPHComponentsAppService
    {
        private readonly ILPHComponentsService _lphComponentsService;

        public LPHComponentsAppService(ILPHComponentsService lphComponentsService) : base(lphComponentsService)
        {
            _lphComponentsService = lphComponentsService;
        }
    }
}
