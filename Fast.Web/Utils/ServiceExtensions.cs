using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using System.Collections.Generic;
using System.Linq;

namespace Fast.Web.Utils
{
	public static class ServiceExtensions
	{
		#region ::Employee Service::
		public static EmployeeModel GetModelByEmpId(this IEmployeeAppService service, string empID)
		{
			EmployeeModel result = new EmployeeModel();
			if (string.IsNullOrEmpty(empID))
				return result;

			string emp = service.GetBy("EmployeeID", empID, true);
			if (!string.IsNullOrEmpty(emp))
				result = emp.DeserializeToEmployee();

			return result;
		}
		#endregion

		#region ::JobTitle Service::
		public static JobTitleModel GetModelById(this IJobTitleAppService service, long jtID)
		{
			JobTitleModel result = new JobTitleModel();
			if (jtID == 0)
				return result;

			string jt = service.GetById(jtID, true);
			if (!string.IsNullOrEmpty(jt))
				result = jt.DeserializeToJobTitle();

			return result;
		}

		public static long AddModel(this IJobTitleAppService service, JobTitleModel model)
		{
			string data = JsonHelper<JobTitleModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void UpdateModel(this IJobTitleAppService service, JobTitleModel model)
		{
			string data = JsonHelper<JobTitleModel>.Serialize(model);

			service.Update(data);
		}
		#endregion

		#region ::User Service::
		public static UserModel GetModelById(this IUserAppService service, long userID)
		{
			UserModel result = new UserModel();
			if (userID == 0)
				return result;

			string user = service.GetById(userID, true);
			if (!string.IsNullOrEmpty(user))
				result = user.DeserializeToUser();

			return result;
		}

		public static long AddModel(this IUserAppService service, UserModel model)
		{
			string data = JsonHelper<UserModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void UpdateModel(this IUserAppService service, UserModel model)
		{
			string data = JsonHelper<UserModel>.Serialize(model);

			service.Update(data);
		}
		#endregion

		#region ::Menu Service::
		public static MenuModel GetModelById(this IMenuAppService service, long menuID)
		{
			MenuModel result = new MenuModel();
			if (menuID == 0)
				return result;

			string user = service.GetById(menuID, true);
			if (!string.IsNullOrEmpty(user))
				result = user.DeserializeToMenu();

			return result;
		}
		#endregion

		#region ::Location Service::
		public static List<long> GetLocIDListByLocType(this ILocationAppService service, long locID, string locationType)
		{
			List<long> result = new List<long>();
			string locations = service.GetAll();
			List<LocationModel> locationModelList = locations.DeserializeToLocationList();

			if (locationType == "country")
			{
				result.Add(locID);
				List<long> prodCenterIdList = locationModelList.Where(x => x.ParentID == locID).Select(x => x.ID).ToList();
				if (prodCenterIdList.Count > 0)
				{
					result.AddRange(prodCenterIdList);
					foreach (var pcID in prodCenterIdList)
					{
						List<long> depIdList = locationModelList.Where(x => x.ParentID == pcID).Select(x => x.ID).ToList();
						if (depIdList.Count > 0)
						{
							result.AddRange(depIdList);
							foreach (var depId in depIdList)
							{
								List<long> subDepIdList = locationModelList.Where(x => x.ParentID == depId).Select(x => x.ID).ToList();
								if (subDepIdList.Count > 0)
								{
									result.AddRange(subDepIdList);
								}
							}
						}
					}
				}
			}
			else if (locationType == "productioncenter")
			{
				result.Add(locID);
				List<long> depIdList = locationModelList.Where(x => x.ParentID == locID).Select(x => x.ID).ToList();
				if (depIdList.Count > 0)
				{
					result.AddRange(depIdList);
					foreach (var depId in depIdList)
					{
						List<long> subDepIdList = locationModelList.Where(x => x.ParentID == depId).Select(x => x.ID).ToList();
						if (subDepIdList.Count > 0)
						{
							result.AddRange(subDepIdList);
						}
					}
				}
			}
			else if (locationType == "department")
			{
				result.Add(locID);
				List<long> subDepIdList = locationModelList.Where(x => x.ParentID == locID).Select(x => x.ID).ToList();
				if (subDepIdList.Count > 0)
				{
					result.AddRange(subDepIdList);
				}
			}
			else
			{
				result.Add(locID);
			}

			return result;
		}
		#endregion

		#region ::LPH Secondary Services::
		public static long AddModel(this ILPHAppService service, LPHModel model)
		{
			string data = JsonHelper<LPHModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static long AddModel(this ILPHValueHistoriesAppService service, LPHValueHistoriesModel model)
		{
			string data = JsonHelper<LPHValueHistoriesModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void UpdateModel(this ILPHAppService service, LPHModel model)
		{
			string data = JsonHelper<LPHModel>.Serialize(model);

			service.Update(data);
		}

		public static long AddModel(this ILPHLocationsAppService service, LPHLocationsModel model)
		{
			string data = JsonHelper<LPHLocationsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void UpdateModel(this ILPHLocationsAppService service, LPHLocationsModel model)
		{
			string data = JsonHelper<LPHLocationsModel>.Serialize(model);

			service.Update(data);
		}

		public static long AddModel(this ILPHComponentsAppService service, LPHComponentsModel model)
		{
			string data = JsonHelper<LPHComponentsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void AddRangeModel(this ILPHComponentsAppService service, List<LPHComponentsModel> models)
		{
			string dataList = JsonHelper<LPHComponentsModel>.Serialize(models);

			service.AddRange(dataList);
		}

		public static void UpdateModel(this ILPHComponentsAppService service, LPHComponentsModel model)
		{
			string data = JsonHelper<LPHComponentsModel>.Serialize(model);

			service.Update(data);
		}

		public static long AddModel(this ILPHApprovalsAppService service, LPHApprovalsModel model)
		{
			string data = JsonHelper<LPHApprovalsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void UpdateModel(this ILPHApprovalsAppService service, LPHApprovalsModel model)
		{
			string data = JsonHelper<LPHApprovalsModel>.Serialize(model);

			service.Update(data);
		}

		public static long AddModel(this ILPHValuesAppService service, LPHValuesModel model)
		{
			string data = JsonHelper<LPHValuesModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void AddRangeModel(this ILPHValuesAppService service, List<LPHValuesModel> models)
		{
			string dataList = JsonHelper<LPHValuesModel>.Serialize(models);

			service.AddRange(dataList);
		}

		public static void UpdateModel(this ILPHValuesAppService service, LPHValuesModel model)
		{
			string data = JsonHelper<LPHValuesModel>.Serialize(model);

			service.Update(data);
		}

		public static long AddModel(this ILPHSubmissionsAppService service, LPHSubmissionsModel model)
		{
			string data = JsonHelper<LPHSubmissionsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void UpdateModel(this ILPHSubmissionsAppService service, LPHSubmissionsModel model)
		{
			string data = JsonHelper<LPHSubmissionsModel>.Serialize(model);

			service.Update(data);
		}

		public static long AddModel(this ILPHExtrasAppService service, LPHExtrasModel model)
		{
			string data = JsonHelper<LPHExtrasModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void UpdateModel(this ILPHExtrasAppService service, LPHExtrasModel model)
		{
			string data = JsonHelper<LPHExtrasModel>.Serialize(model);

			service.Update(data);
		}
		#endregion

		#region ::LPH Primary Services::
		public static void UpdateModel(this IPPLPHValuesAppService service, PPLPHValuesModel model)
		{
			string data = JsonHelper<PPLPHValuesModel>.Serialize(model);

			service.Update(data);
		}

		public static void UpdateModel(this IPPLPHAppService service, PPLPHModel model)
		{
			string data = JsonHelper<PPLPHModel>.Serialize(model);

			service.Update(data);
		}

		public static long AddModel(this IPPLPHAppService service, PPLPHModel model)
		{
			string data = JsonHelper<PPLPHModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static long AddModel(this IPPLPHLocationsAppService service, PPLPHLocationsModel model)
		{
			string data = JsonHelper<PPLPHLocationsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static long AddModel(this IPPLPHComponentsAppService service, PPLPHComponentsModel model)
		{
			string data = JsonHelper<PPLPHComponentsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void AddRangeModel(this IPPLPHComponentsAppService service, List<PPLPHComponentsModel> models)
		{
			string dataList = JsonHelper<PPLPHComponentsModel>.Serialize(models);

			service.AddRange(dataList);
		}


		public static long AddModel(this IPPLPHApprovalsAppService service, PPLPHApprovalsModel model)
		{
			string data = JsonHelper<PPLPHApprovalsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static long AddModel(this IPPLPHValuesAppService service, PPLPHValuesModel model)
		{
			string data = JsonHelper<PPLPHValuesModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static void AddRangeModel(this IPPLPHValuesAppService service, List<PPLPHValuesModel> models)
		{
			string dataList = JsonHelper<PPLPHValuesModel>.Serialize(models);

			service.AddRange(dataList);
		}

		public static long AddModel(this IPPLPHSubmissionsAppService service, PPLPHSubmissionsModel model)
		{
			string data = JsonHelper<PPLPHSubmissionsModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}

		public static long AddModel(this IPPLPHExtrasAppService service, PPLPHExtrasModel model)
		{
			string data = JsonHelper<PPLPHExtrasModel>.Serialize(model);

			long result = service.Add(data);

			return result;
		}
		#endregion
	}
}