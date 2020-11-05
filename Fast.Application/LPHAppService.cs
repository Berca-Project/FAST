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
    public class LPHAppService: AppServiceBase<LPH>, ILPHAppService
    {
        private readonly ILPHService _lphService;
        public LPHAppService(ILPHService lphService) : base(lphService)
        {
            _lphService = lphService;
        }
        public LPH GetObjectByID(long id)
        {
            LPH lph = _lphService.GetById(id);
            return lph == null ? null : lph;
        }
    }
}
