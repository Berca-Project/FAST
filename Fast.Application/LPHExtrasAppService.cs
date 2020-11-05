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
    public class LPHExtrasAppService:AppServiceBase<LPHExtras>, ILPHExtrasAppService
    {
        private readonly ILPHExtrasService _lphExtrasService;
        public LPHExtrasAppService(ILPHExtrasService lphExtrasService) : base(lphExtrasService)
        {
            _lphExtrasService = lphExtrasService;
        }
    }
}
