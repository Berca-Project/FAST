using Fast.Application;
using Fast.Application.Interfaces;
using Fast.Domain.Entities;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class MenuAppService : AppServiceBase<Menu>, IMenuAppService
	{
		public MenuAppService(IMenuServiceService serviceBase) : base(serviceBase)
		{
		}
	}
}