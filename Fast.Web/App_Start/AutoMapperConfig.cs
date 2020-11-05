using AutoMapper;
using Fast.Web.Models;

namespace Fast.Web.App_Start
{
	public class AutomapperConfig : Profile
	{
		public AutomapperConfig()
		{
			CreateMap<UserModel, UserModel>();			
		}
	}
}