using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Models.LPH;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Fast.Web.Utils
{
	public static class LPHSPHelper
	{
		#region ::Setup LPH Approval Model::
		public static LPHEditModel SetupLPHApprovalModel(
			ILPHAppService _lphAppService,
			ILPHExtrasAppService _lphExtrasAppService,
			ILPHComponentsAppService _lphComponentsAppService,
			ILPHValuesAppService _lphValuesAppService, long submitid,
			long lphID)
		{
			var model = new LPHEditModel();
			string lph = _lphAppService.GetById(lphID, true);
			LPHModel lphModel = lph.DeserializeToLPH();
			model.LPH = lphModel;

			string extra = _lphExtrasAppService.FindBy("LPHID", lphModel.ID, true);
			List<LPHExtrasModel> extraList = extra.DeserializeToLPHExtraList();

			for (int i = 0; i < extraList.Count; i++)
			{
				extraList[i].Value = extraList[i].Value != null ? Uri.EscapeDataString(extraList[i].Value) : "";
			}

			model.Extras = extraList.OrderBy(x => x.ID).ToList();

			string compo = _lphComponentsAppService.FindBy("LPHID", lphModel.ID, true);
			IEnumerable<LPHComponentsModel> compoList = compo.DeserializeToLPHComponentList().OrderBy(x => x.ID).ToList();

			model.CompoVal = new List<LPHCompoValModel>();
            string values = _lphValuesAppService.FindBy("SubmissionID", submitid, true);
			List<LPHValuesModel> valueModelList = values.DeserializeToLPHValueList();

			foreach (var component in compoList)
			{
				var valueModel = valueModelList.Where(x => x.LPHComponentID == component.ID).FirstOrDefault();
				if (valueModel == null)
					valueModel = _lphValuesAppService.GetBy("LPHComponentID", component.ID, true).DeserializeToLPHValue();

				var CompoVal = new LPHCompoValModel();
				CompoVal.Component = component;
				CompoVal.Value = valueModel;
				CompoVal.Value.Value = CompoVal.Value.Value != null ? Uri.EscapeDataString(CompoVal.Value.Value) : "";

				model.CompoVal.Add(CompoVal);
			}

			return model;
		}
		#endregion

		#region ::Is Redirect To Edit::
		public static bool IsRedirectToEdit(
			ILPHAppService _lphAppService,
			ILPHSubmissionsAppService _lphSubmissionsAppService,
			string menuTitle,
			string accountName,
			out long resultID)
		{
			List<QueryFilter> lphFilter = new List<QueryFilter>();
			lphFilter.Add(new QueryFilter("MenuTitle", menuTitle));
			lphFilter.Add(new QueryFilter("ModifiedBy", accountName));
			lphFilter.Add(new QueryFilter("IsDeleted", "0"));

			string checkDatas = _lphAppService.Find(lphFilter);
			LPHModel result = checkDatas.DeserializeToLPHList().OrderBy(x => x.ID).LastOrDefault();
			if (result != null)
			{
				string checkSubmit = _lphSubmissionsAppService.GetBy("LPHID", result.ID, false);
				LPHSubmissionsModel submitModel = checkSubmit.DeserializeToLPHSubmissions();
				if (submitModel.IsDeleted)
				{
					resultID = result.ID;
					return true;
				}
			}

			resultID = 0;
			return false;
		}

		public static bool IsRedirectToEdit2(
			ILPHAppService _lphAppService,
			ILPHSubmissionsAppService _lphSubmissionsAppService,
			string menuTitle,
			string accountName,
			string Shift,
			out LPHSubmissionsModel resultID,
			out int oneDay)
		{
			List<QueryFilter> lphFilter = new List<QueryFilter>();
			lphFilter.Add(new QueryFilter("MenuTitle", menuTitle));
			lphFilter.Add(new QueryFilter("ModifiedBy", accountName));
			lphFilter.Add(new QueryFilter("IsDeleted", "0"));

			string checkDatas = _lphAppService.Find(lphFilter);
			LPHModel result = checkDatas.DeserializeToLPHList().OrderBy(x => x.ID).LastOrDefault();
			if (result != null)
			{
				string checkSubmit = _lphSubmissionsAppService.GetBy("LPHID", result.ID, false);
				LPHSubmissionsModel submitModel = checkSubmit.DeserializeToLPHSubmissions();
				if (submitModel.IsDeleted)
				{
					DateTime submitDate = submitModel.ModifiedDate ?? DateTime.Now;
					if (submitDate.ToString("yyyy-MM-dd") == DateTime.Now.ToString("yyyy-MM-dd"))
					{
						if (submitModel.Shift.Trim() != Shift.Trim())
						{
							resultID = submitModel;
							oneDay = 0;
							return true;
						}
						else
						{
							resultID = submitModel;
							oneDay = 1;
							return false;
						}
					}
					else
					{
						resultID = submitModel;
						oneDay = 0;
						return true;
					}
				}
			}

			resultID = null;
			oneDay = 0;
			return false;
		}
		#endregion

		#region ::Get Machine List::
		public static List<MachineModel> GetMachineList(
			IMachineAppService _machineAppService,
			List<long> locationIdList,
			List<string> modelMachineList,
			bool isMaker = false,
			bool isPacker = false, string other = "", bool noLOC = false)
		{
			string machine = _machineAppService.GetAll(true);
			List<MachineModel> machineList = machine.DeserializeToMachineList();

			if (isMaker)
				machineList = machineList.Where(x => x.Location.Contains("MK")).OrderBy(x => x.Code).ToList();
			else if (isPacker)
				machineList = machineList.Where(x => x.Location.Contains("PC")).OrderBy(x => x.Code).ToList();
			else
			{
				if (other != "")
				{
					if (other == "XXX") //all machine
										//machineList = machineList.Where(x => x.Location.Contains("PT") || x.Location.Contains("CB") || x.Location.Contains("RW") || x.Location.Contains("SD") || x.Location.Contains("SZ") || x.Location.Contains("CT")).OrderBy(x => x.Code).ToList();
						machineList = machineList.Where(x => x.Code != null).OrderBy(x => x.Code).ToList();
					else if (other == "ALL") //all machine rata smua
						machineList = machineList.Where(x => x.Code != null).OrderBy(x => x.Code).ToList();
					else
						machineList = machineList.Where(x => x.Location.Contains(other)).OrderBy(x => x.Code).ToList();
				}
				else //mesin selain maker packer
					machineList = machineList.Where(x => !x.Location.Contains("PC") && !x.Location.Contains("MK")).OrderBy(x => x.Code).ToList();
			}

			//if (modelMachineList.Count > 0)
			//{
			//	machineList = machineList.Where(x => modelMachineList.Any(z => z == x.Code)).OrderBy(x => x.Code).ToList();
			//}

            if(!noLOC)
			    machineList = machineList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
			
            machineList = machineList.GroupBy(x => x.Code).Select(x => x.FirstOrDefault()).ToList();

			return machineList;
		}
		#endregion

		#region ::Get Machine Type List::
		public static List<MachineModel> GetMachineTypeList(
			IMachineAppService _machineAppService,
			IReferenceAppService _referenceAppService,
			List<long> locationIdList,
			List<string> modelMachineList, bool noLOC = false)
		{
			string machines = _machineAppService.GetAll(true);
			List<MachineModel> machineModelList = machines.DeserializeToMachineList();
			if (modelMachineList.Count < 1)
				machineModelList = machineModelList.Where(x => !x.Code.StartsWith("P") && !x.Code.StartsWith("M") && x.Code.Length > 1).OrderBy(x => x.Code).ToList();
			else
				machineModelList = machineModelList.Where(x => !x.Code.StartsWith("P") && !x.Code.StartsWith("M") && x.Code.Length > 1 && modelMachineList.Any(z => z == x.Code)).OrderBy(x => x.Code).ToList();
            
            if(!noLOC)
			    machineModelList = machineModelList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
			
            machineModelList = machineModelList.GroupBy(x => x.Code).Select(x => x.FirstOrDefault()).ToList();

			foreach (var item in machineModelList)
			{
				string machineType = _referenceAppService.GetDetailById(item.MachineTypeID, true);
				ReferenceDetailModel machineTypeModel = machineType.DeserializeToRefDetail();

				item.MachineType = machineTypeModel.Code;
			}

			machineModelList = machineModelList.OrderBy(x => x.Location).ToList();
			return machineModelList;
		}
		#endregion

		#region ::Get Brand List::
		public static List<ReferenceDetailModel> GetBrandList(IBrandAppService _brandAppService, List<long> locationIdList, List<string> modelBrandList, bool isFAcode = false, int nonFA = 0, bool noLOC = false)
		{
			string brand = _brandAppService.GetAll();
			List<BrandModel> brandList = brand.DeserializeToBrandList();
            if(!noLOC)
            {
                if (isFAcode)
                    brandList = brandList.Where(x => x.Code.StartsWith("F") && locationIdList.Any(y => y == x.LocationID)).ToList();
                else
                {
                    if (nonFA > 0)
                        brandList = brandList.Where(x => !x.Code.StartsWith("F") && locationIdList.Any(y => y == x.LocationID)).ToList();
                    else
                        brandList = brandList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
                }
            }
            else //ignore location
            {
                if (isFAcode)
                    brandList = brandList.Where(x => x.Code.StartsWith("F")).ToList();
                else
                {
                    if (nonFA > 0)
                        brandList = brandList.Where(x => !x.Code.StartsWith("F")).ToList();
                    //else
                    //    brandList = brandList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();
                }

            }

            List<ReferenceDetailModel> brandModelList = new List<ReferenceDetailModel>();
			brandModelList = brandList.Select(x => new ReferenceDetailModel() { Code = x.Code, Description = x.Description }).ToList();

			//if (modelBrandList.Count > 0)
			//{
			//	brandModelList = brandModelList.Where(x => x.Code != null && modelBrandList.Any(z => z == x.Code)).OrderBy(x => x.Code).ToList();
			//}

			return brandModelList;
		}
		#endregion

		#region ::Get Extra Role List::
		public static List<EmployeeRoleModel> GetExtraRole(IUserRoleAppService _userRoleAppService, IUserAppService _userAppService, List<EmployeeModel> employeeList)
		{
			string userroles = _userRoleAppService.GetAll();
			List<UserRoleModel> userRoleModelList = userroles.DeserializeToUserRoleList();
			List<EmployeeRoleModel> extraEmployeeRoles = new List<EmployeeRoleModel>();

			foreach (var item in userRoleModelList)
			{
				string userStr = _userAppService.GetById(item.UserID);
				UserModel userModel = userStr.DeserializeToUser();
				EmployeeModel emp = employeeList.Where(x => x.EmployeeID == userModel.EmployeeID).FirstOrDefault();
				if (emp != null)
				{
					extraEmployeeRoles.Add(new EmployeeRoleModel
					{
						RoleName = item.RoleName.ToLower(),
						EmployeeModel = emp
					});
				}
			}

			return extraEmployeeRoles;
		}
		#endregion

		#region ::Get Electrician List::
		public static List<EmployeeModel> GetElectricianList(List<EmployeeModel> employeeList, string modelPosition, List<string> positionList, List<UserModel> userList, List<long> locationIdList, List<EmployeeRoleModel> extraEmployeeRoles = null, bool noLOC = false)
		{
			List<EmployeeModel> empList = new List<EmployeeModel>();

			if (string.IsNullOrEmpty(modelPosition))
			{
				if (!positionList.Contains("Electrician")) positionList.Add("Electrician");
				empList = employeeList.Where(x => x.PositionDesc != null && positionList.Any(y => y.Equals(x.PositionDesc, StringComparison.OrdinalIgnoreCase))).OrderBy(x => x.FullName).ToList();
			}
			else
			{
				empList = employeeList.Where(x => x.EmployeeID != null && x.EmployeeID.Equals(modelPosition)).OrderBy(x => x.FullName).ToList();
			}

			if (extraEmployeeRoles != null)
			{
				List<EmployeeRoleModel> empExtraList = extraEmployeeRoles.Where(x => x.RoleName.Contains("electrician")).ToList();
				if (empExtraList.Count > 0)
				{
					foreach (var item in empExtraList)
					{
						empList.Add(item.EmployeeModel);
					}
				}
			}

            if (!noLOC)
                empList = FilterEmpListByLocation(userList, locationIdList, empList);
            else
                empList = FilterEmpListByLocation(userList, locationIdList, empList, true);

            return empList;
		}
		#endregion

		#region ::Get Prod Tech List::
		public static List<EmployeeModel> GetProdTechList(List<EmployeeModel> employeeList, string modelPosition, List<string> positionList, List<UserModel> userList, List<long> locationIdList, List<EmployeeRoleModel> extraEmployeeRoles = null, bool noLOC = false)
		{
			List<EmployeeModel> empList = new List<EmployeeModel>();

			if (string.IsNullOrEmpty(modelPosition))
			{
				if (!positionList.Contains("Production Technician")) positionList.Add("Production Technician");
				empList = employeeList.Where(x => x.PositionDesc != null && positionList.Any(y => y.Equals(x.PositionDesc, StringComparison.OrdinalIgnoreCase))).OrderBy(x => x.FullName).ToList();
			}
			else
			{
				empList = employeeList.Where(x => x.EmployeeID != null && x.EmployeeID.Equals(modelPosition)).OrderBy(x => x.FullName).ToList();
			}

			if (extraEmployeeRoles != null)
			{
				List<EmployeeRoleModel> empExtraList = extraEmployeeRoles.Where(x => x.RoleName.Contains("prodtech")).ToList();
				if (empExtraList.Count > 0)
				{
					foreach (var item in empExtraList)
					{
						empList.Add(item.EmployeeModel);
					}
				}
			}

            if (!noLOC)
                empList = FilterEmpListByLocation(userList, locationIdList, empList);
            else
                empList = FilterEmpListByLocation(userList, locationIdList, empList, true);

            return empList;
		}
		#endregion

		#region ::Get Foreman / Supervisor List::
		public static List<EmployeeModel> GetForemanList(List<EmployeeModel> employeeList, string modelPosition, List<UserModel> userList, List<long> locationIdList, List<EmployeeRoleModel> extraEmployeeRoles = null, bool noLOC = false)
		{
			List<EmployeeModel> empList = new List<EmployeeModel>();

			if (string.IsNullOrEmpty(modelPosition))
			{
				empList = employeeList.Where(x => x.PositionDesc != null && x.PositionDesc.Contains("Foreman")).OrderBy(x => x.FullName).ToList();
			}
			else
			{
				empList = employeeList.Where(x => x.EmployeeID != null && x.EmployeeID.Equals(modelPosition)).OrderBy(x => x.FullName).ToList();
			}

			if (extraEmployeeRoles != null)
			{
				List<EmployeeRoleModel> empExtraList = extraEmployeeRoles.Where(x => x.RoleName.Contains("foreman")).ToList();
				if (empExtraList.Count > 0)
				{
					foreach (var item in empExtraList)
					{
						empList.Add(item.EmployeeModel);
					}
				}
			}

            if (!noLOC)
                empList = FilterEmpListByLocation(userList, locationIdList, empList);
            else
                empList = FilterEmpListByLocation(userList, locationIdList, empList, true);

            return empList;
		}
		#endregion

		#region ::Get Mechanic List::
		public static List<EmployeeModel> GetMechanicList(List<EmployeeModel> employeeList, string modelPosition, List<string> positionList, List<UserModel> userList, List<long> locationIdList, List<EmployeeRoleModel> extraEmployeeRoles = null, bool noLOC = false)
		{
			List<EmployeeModel> empList = new List<EmployeeModel>();

			if (string.IsNullOrEmpty(modelPosition))
			{
				if (!positionList.Contains("Mechanic")) positionList.Add("Mechanic");
				empList = employeeList.Where(x => x.PositionDesc != null && positionList.Any(y => y.Equals(x.PositionDesc, StringComparison.OrdinalIgnoreCase))).OrderBy(x => x.FullName).ToList();
			}
			else
			{
				empList = employeeList.Where(x => x.EmployeeID != null && x.EmployeeID.Equals(modelPosition)).OrderBy(x => x.FullName).ToList();
			}

			if (extraEmployeeRoles != null)
			{
				List<EmployeeRoleModel> empExtraList = extraEmployeeRoles.Where(x => x.RoleName.Contains("mechanic")).ToList();
				if (empExtraList.Count > 0)
				{
					foreach (var item in empExtraList)
					{
						empList.Add(item.EmployeeModel);
					}
				}
			}

            if (!noLOC)
                empList = FilterEmpListByLocation(userList, locationIdList, empList);
            else
                empList = FilterEmpListByLocation(userList, locationIdList, empList, true);

            return empList;
		}
		#endregion

		#region ::Get Team Lead / Supervisor List::
		public static List<EmployeeModel> GetTeamLeadList(List<EmployeeModel> employeeList, string modelPosition, List<string> positionList, List<UserModel> userList, List<long> locationIdList, List<EmployeeRoleModel> extraEmployeeRoles = null, bool noLOC = false)
		{
			List<EmployeeModel> empList = new List<EmployeeModel>();

			if (string.IsNullOrEmpty(modelPosition))
			{
				if (!positionList.Contains("Supervisor")) positionList.Add("Supervisor");
				if (!positionList.Contains("Team Lead")) positionList.Add("Team Lead");
				if (!positionList.Contains("Foreman")) positionList.Add("Foreman");
				if (!positionList.Contains("Lead")) positionList.Add("Lead");

				empList = employeeList.Where(x => x.PositionDesc != null && positionList.Any(y => y.Equals(x.PositionDesc, StringComparison.OrdinalIgnoreCase))).OrderBy(x => x.FullName).ToList();
			}
			else
			{
				empList = employeeList.Where(x => x.EmployeeID != null && x.EmployeeID.Equals(modelPosition)).OrderBy(x => x.FullName).ToList();
			}


			if (extraEmployeeRoles != null)
			{
				List<EmployeeRoleModel> empExtraList = extraEmployeeRoles.Where(x => x.RoleName.Contains("supervisor")).ToList();
				if (empExtraList.Count > 0)
				{
					foreach (var item in empExtraList)
					{
						empList.Add(item.EmployeeModel);
					}
				}
			}

            if (!noLOC) //false then noLoc filter
                empList = FilterEmpListByLocation(userList, locationIdList, empList);
            else
                empList = FilterEmpListByLocation(userList, locationIdList, empList, true);

            return empList;
		}
		#endregion

		#region ::Get General Worker List::
		public static List<EmployeeModel> GetGeneralWorkerList(List<EmployeeModel> employeeList, string modelPosition, List<string> positionList, List<UserModel> userList, List<long> locationIdList, List<EmployeeRoleModel> extraEmployeeRoles = null, bool noLOC = false)
		{
			List<EmployeeModel> empList = new List<EmployeeModel>();

			if (string.IsNullOrEmpty(modelPosition))
			{
				if (!positionList.Contains("General Worker")) positionList.Add("General Worker");
				empList = employeeList.Where(x => x.PositionDesc != null && positionList.Any(y => y.Equals(x.PositionDesc, StringComparison.OrdinalIgnoreCase))).OrderBy(x => x.FullName).ToList();
			}
			else
			{
				empList = employeeList.Where(x => x.EmployeeID != null && x.EmployeeID.Equals(modelPosition)).OrderBy(x => x.FullName).ToList();
			}

			if (extraEmployeeRoles != null)
			{
				List<EmployeeRoleModel> empExtraList = extraEmployeeRoles.Where(x => x.RoleName.Contains("gw") || x.RoleName.Contains("general")).ToList();
				if (empExtraList.Count > 0)
				{
					foreach (var item in empExtraList)
					{
						empList.Add(item.EmployeeModel);
					}
				}
			}

            if(!noLOC)
			    empList = FilterEmpListByLocation(userList, locationIdList, empList);
            else
                empList = FilterEmpListByLocation(userList, locationIdList, empList ,true);

            return empList;
		}
		#endregion

		#region ::Filter Emp List by Location::
		private static List<EmployeeModel> FilterEmpListByLocation(List<UserModel> userList, List<long> locationIdList, List<EmployeeModel> empList, bool noLOC=false)
		{

			List<UserModel> userFilteredList = new List<UserModel>();

			List<string> empIDList = empList.Select(x => x.EmployeeID).ToList();

			userFilteredList = userList.Where(x => empIDList.Any(y => y == x.EmployeeID) && x.IsActive && x.IsFast).ToList();

            if(!noLOC) //false then filter by location
			    userFilteredList = userFilteredList.Where(x => locationIdList.Any(y => y == x.LocationID)).ToList();

			List<string> userEmpIDList = userFilteredList.Select(x => x.EmployeeID).ToList();

			empList = empList.Where(x => userEmpIDList.Any(y => y == x.EmployeeID)).ToList();

			return empList;
		}
		#endregion
	}
}