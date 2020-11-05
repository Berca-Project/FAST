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
    public class PPLPHLocationsAppService : AppServiceBase<PPLPHLocations>, IPPLPHLocationsAppService
    {
        private readonly IPPLPHLocationsService _PPLphLocationsService;

        public PPLPHLocationsAppService(IPPLPHLocationsService PPLphLocationsService) : base(PPLphLocationsService)
        {
            _PPLphLocationsService = PPLphLocationsService;
        }
    }
}
