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
    public class LPHLocationsAppService: AppServiceBase<LPHLocations>, ILPHLocationsAppService
    {
        private readonly ILPHLocationsService _lphLocationsService;

        public LPHLocationsAppService(ILPHLocationsService lphLocationsService) : base(lphLocationsService)
        {
            _lphLocationsService = lphLocationsService;
        }
    }
}
