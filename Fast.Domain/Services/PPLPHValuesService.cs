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
    public class PPLPHValuesService : ServiceBase<PPLPHValues>, IPPLPHValuesService
    {
        public PPLPHValuesService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}
