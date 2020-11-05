using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Repositories;
using Fast.Domain.Interfaces.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Domain.Services
{
    public class PPLPHService : ServiceBase<PPLPH>, IPPLPHService
    {
        public PPLPHService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
