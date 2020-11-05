using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Models.Report;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Fast.Web
{
	public static class StringExtensions
	{
		public static WppPrimaryModel DeserializeToWppPrimary(this string value)
		{
			return string.IsNullOrEmpty(value) ? new WppPrimaryModel() : JsonConvert.DeserializeObject<WppPrimaryModel>(value);
		}

		public static List<WppPrimaryModel> DeserializeToWppPrimaryList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<WppPrimaryModel>() : JsonConvert.DeserializeObject<List<WppPrimaryModel>>(value);
		}

		public static MachineAllocationModel DeserializeToMachineAllocation(this string value)
		{
			return string.IsNullOrEmpty(value) ? new MachineAllocationModel() : JsonConvert.DeserializeObject<MachineAllocationModel>(value);
		}

		public static List<MachineAllocationModel> DeserializeToMachineAllocationList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<MachineAllocationModel>() : JsonConvert.DeserializeObject<List<MachineAllocationModel>>(value);
		}

		public static ShuttleRequestModel DeserializeToShuttleRequest(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ShuttleRequestModel() : JsonConvert.DeserializeObject<ShuttleRequestModel>(value);
		}

		public static List<ShuttleRequestModel> DeserializeToShuttleRequestList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ShuttleRequestModel>() : JsonConvert.DeserializeObject<List<ShuttleRequestModel>>(value);
		}

		public static MealRequestModel DeserializeToMealRequest(this string value)
		{
			return string.IsNullOrEmpty(value) ? new MealRequestModel() : JsonConvert.DeserializeObject<MealRequestModel>(value);
		}

		public static List<MealRequestModel> DeserializeToMealRequestList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<MealRequestModel>() : JsonConvert.DeserializeObject<List<MealRequestModel>>(value);
		}

		public static TrainingTitleMachineTypeModel DeserializeToTrainingTitleMachineType(this string value)
		{
			return string.IsNullOrEmpty(value) ? new TrainingTitleMachineTypeModel() : JsonConvert.DeserializeObject<TrainingTitleMachineTypeModel>(value);
		}

		public static List<TrainingTitleMachineTypeModel> DeserializeToTrainingTitleMachineTypeList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<TrainingTitleMachineTypeModel>() : JsonConvert.DeserializeObject<List<TrainingTitleMachineTypeModel>>(value);
		}

		public static TrainingTitleModel DeserializeToTrainingTitle(this string value)
		{
			return string.IsNullOrEmpty(value) ? new TrainingTitleModel() : JsonConvert.DeserializeObject<TrainingTitleModel>(value);
		}

		public static List<TrainingTitleModel> DeserializeToTrainingTitleList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<TrainingTitleModel>() : JsonConvert.DeserializeObject<List<TrainingTitleModel>>(value);
		}

		public static MaterialCodeModel DeserializeToMaterialCode(this string value)
		{
			return string.IsNullOrEmpty(value) ? new MaterialCodeModel() : JsonConvert.DeserializeObject<MaterialCodeModel>(value);
		}

		public static List<MaterialCodeModel> DeserializeToMaterialCodeList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<MaterialCodeModel>() : JsonConvert.DeserializeObject<List<MaterialCodeModel>>(value);
		}

		public static UserRoleModel DeserializeToUserRole(this string value)
		{
			return string.IsNullOrEmpty(value) ? new UserRoleModel() : JsonConvert.DeserializeObject<UserRoleModel>(value);
		}

		public static List<UserRoleModel> DeserializeToUserRoleList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<UserRoleModel>() : JsonConvert.DeserializeObject<List<UserRoleModel>>(value);
		}

		public static BrandModel DeserializeToBrand(this string value)
		{
			return string.IsNullOrEmpty(value) ? new BrandModel() : JsonConvert.DeserializeObject<BrandModel>(value);
		}

		public static List<BrandModel> DeserializeToBrandList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<BrandModel>() : JsonConvert.DeserializeObject<List<BrandModel>>(value);
		}

		public static BrandConversionModel DeserializeToBrandConversion(this string value)
		{
			return string.IsNullOrEmpty(value) ? new BrandConversionModel() : JsonConvert.DeserializeObject<BrandConversionModel>(value);
		}

		public static List<BrandConversionModel> DeserializeToBrandConversionList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<BrandConversionModel>() : JsonConvert.DeserializeObject<List<BrandConversionModel>>(value);
		}

		public static BlendModel DeserializeToBlend(this string value)
		{
			return string.IsNullOrEmpty(value) ? new BlendModel() : JsonConvert.DeserializeObject<BlendModel>(value);
		}

		public static List<BlendModel> DeserializeToBlendList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<BlendModel>() : JsonConvert.DeserializeObject<List<BlendModel>>(value);
		}

		public static ManPowerModel DeserializeToManPower(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ManPowerModel() : JsonConvert.DeserializeObject<ManPowerModel>(value);
		}

		public static List<ManPowerModel> DeserializeToManPowerList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ManPowerModel>() : JsonConvert.DeserializeObject<List<ManPowerModel>>(value);
		}

		public static CalendarModel DeserializeToCalendar(this string value)
		{
			return string.IsNullOrEmpty(value) ? new CalendarModel() : JsonConvert.DeserializeObject<CalendarModel>(value);
		}

		public static List<CalendarModel> DeserializeToCalendarList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<CalendarModel>() : JsonConvert.DeserializeObject<List<CalendarModel>>(value);
		}

		public static CalendarHolidayModel DeserializeToCalendarHoliday(this string value)
		{
			return string.IsNullOrEmpty(value) ? new CalendarHolidayModel() : JsonConvert.DeserializeObject<CalendarHolidayModel>(value);
		}

		public static List<CalendarHolidayModel> DeserializeToCalendarHolidayList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<CalendarHolidayModel>() : JsonConvert.DeserializeObject<List<CalendarHolidayModel>>(value);
		}

		public static EmployeeModel DeserializeToEmployee(this string value)
		{
			return string.IsNullOrEmpty(value) ? new EmployeeModel() : JsonConvert.DeserializeObject<EmployeeModel>(value);
		}

		public static List<EmployeeModel> DeserializeToEmployeeList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<EmployeeModel>() : JsonConvert.DeserializeObject<List<EmployeeModel>>(value);
		}

		public static SimpleEmployeeModel DeserializeToSimpleEmployee(this string value)
		{
			return string.IsNullOrEmpty(value) ? new SimpleEmployeeModel() : JsonConvert.DeserializeObject<SimpleEmployeeModel>(value);
		}

		public static List<SimpleEmployeeModel> DeserializeToSimpleEmployeeList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<SimpleEmployeeModel>() : JsonConvert.DeserializeObject<List<SimpleEmployeeModel>>(value);
		}

		public static EmployeeAllModel DeserializeToEmployeeAll(this string value)
		{
			return string.IsNullOrEmpty(value) ? new EmployeeAllModel() : JsonConvert.DeserializeObject<EmployeeAllModel>(value);
		}

		public static List<EmployeeAllModel> DeserializeToEmployeeAllList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<EmployeeAllModel>() : JsonConvert.DeserializeObject<List<EmployeeAllModel>>(value);
		}

		public static EmployeeOvertimeModel DeserializeToEmployeeOvertime(this string value)
		{
			return string.IsNullOrEmpty(value) ? new EmployeeOvertimeModel() : JsonConvert.DeserializeObject<EmployeeOvertimeModel>(value);
		}

		public static List<EmployeeOvertimeModel> DeserializeToEmployeeOvertimeList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<EmployeeOvertimeModel>() : JsonConvert.DeserializeObject<List<EmployeeOvertimeModel>>(value);
		}

		public static EmployeeLeaveModel DeserializeToEmployeeLeave(this string value)
		{
			return string.IsNullOrEmpty(value) ? new EmployeeLeaveModel() : JsonConvert.DeserializeObject<EmployeeLeaveModel>(value);
		}

		public static List<EmployeeLeaveModel> DeserializeToEmployeeLeaveList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<EmployeeLeaveModel>() : JsonConvert.DeserializeObject<List<EmployeeLeaveModel>>(value);
		}

		public static AccessRightDBModel DeserializeToAccessRight(this string value)
		{
			return string.IsNullOrEmpty(value) ? new AccessRightDBModel() : JsonConvert.DeserializeObject<AccessRightDBModel>(value);
		}

		public static List<AccessRightDBModel> DeserializeToAccessRightList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<AccessRightDBModel>() : JsonConvert.DeserializeObject<List<AccessRightDBModel>>(value);
		}

		public static MenuModel DeserializeToMenu(this string value)
		{
			return string.IsNullOrEmpty(value) ? new MenuModel() : JsonConvert.DeserializeObject<MenuModel>(value);
		}

		public static List<MenuModel> DeserializeToMenuList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<MenuModel>() : JsonConvert.DeserializeObject<List<MenuModel>>(value);
		}

		public static List<ParentMenuModel> DeserializeToParentMenuList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ParentMenuModel>() : JsonConvert.DeserializeObject<List<ParentMenuModel>>(value);
		}

		public static List<ChildMenuModel> DeserializeToChildMenuList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChildMenuModel>() : JsonConvert.DeserializeObject<List<ChildMenuModel>>(value);
		}

		public static RoleModel DeserializeToRole(this string value)
		{
			return string.IsNullOrEmpty(value) ? new RoleModel() : JsonConvert.DeserializeObject<RoleModel>(value);
		}

		public static List<RoleModel> DeserializeToRoleList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<RoleModel>() : JsonConvert.DeserializeObject<List<RoleModel>>(value);
		}

		public static LocationModel DeserializeToLocation(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LocationModel() : JsonConvert.DeserializeObject<LocationModel>(value);
		}

		public static List<LocationModel> DeserializeToLocationList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LocationModel>() : JsonConvert.DeserializeObject<List<LocationModel>>(value);
		}

		public static UserModel DeserializeToUser(this string value)
		{
			return string.IsNullOrEmpty(value) ? new UserModel() : JsonConvert.DeserializeObject<UserModel>(value);
		}

		public static List<UserModel> DeserializeToUserList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<UserModel>() : JsonConvert.DeserializeObject<List<UserModel>>(value);
		}

		public static MachineModel DeserializeToMachine(this string value)
		{
			return string.IsNullOrEmpty(value) ? new MachineModel() : JsonConvert.DeserializeObject<MachineModel>(value);
		}

		public static List<MachineModel> DeserializeToMachineList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<MachineModel>() : JsonConvert.DeserializeObject<List<MachineModel>>(value);
		}

		public static JobTitleModel DeserializeToJobTitle(this string value)
		{
			return string.IsNullOrEmpty(value) ? new JobTitleModel() : JsonConvert.DeserializeObject<JobTitleModel>(value);
		}

		public static List<JobTitleModel> DeserializeToJobTitleList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<JobTitleModel>() : JsonConvert.DeserializeObject<List<JobTitleModel>>(value);
		}

		public static LocationMachineTypeModel DeserializeToLocationMachineType(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LocationMachineTypeModel() : JsonConvert.DeserializeObject<LocationMachineTypeModel>(value);
		}

		public static List<LocationMachineTypeModel> DeserializeToLocationMachineTypeList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LocationMachineTypeModel>() : JsonConvert.DeserializeObject<List<LocationMachineTypeModel>>(value);
		}

		public static UserMachineTypeModel DeserializeToUserMachineType(this string value)
		{
			return string.IsNullOrEmpty(value) ? new UserMachineTypeModel() : JsonConvert.DeserializeObject<UserMachineTypeModel>(value);
		}

		public static List<UserMachineTypeModel> DeserializeToUserMachineTypeList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<UserMachineTypeModel>() : JsonConvert.DeserializeObject<List<UserMachineTypeModel>>(value);
		}

		public static UserMachineModel DeserializeToUserMachine(this string value)
		{
			return string.IsNullOrEmpty(value) ? new UserMachineModel() : JsonConvert.DeserializeObject<UserMachineModel>(value);
		}

		public static List<UserMachineModel> DeserializeToUserMachineList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<UserMachineModel>() : JsonConvert.DeserializeObject<List<UserMachineModel>>(value);
		}

		public static WppPrimModel DeserializeToWppPrim(this string value)
		{
			return string.IsNullOrEmpty(value) ? new WppPrimModel() : JsonConvert.DeserializeObject<WppPrimModel>(value);
		}

		public static List<WppPrimModel> DeserializeToWppPrimList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<WppPrimModel>() : JsonConvert.DeserializeObject<List<WppPrimModel>>(value);
		}

		public static WppModel DeserializeToWpp(this string value)
		{
			return string.IsNullOrEmpty(value) ? new WppModel() : JsonConvert.DeserializeObject<WppModel>(value);
		}

		public static List<WppModel> DeserializeToWppList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<WppModel>() : JsonConvert.DeserializeObject<List<WppModel>>(value);
		}

		public static WppStpModel DeserializeToWppStp(this string value)
		{
			return string.IsNullOrEmpty(value) ? new WppStpModel() : JsonConvert.DeserializeObject<WppStpModel>(value);
		}

		public static List<WppStpModel> DeserializeToWppStpList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<WppStpModel>() : JsonConvert.DeserializeObject<List<WppStpModel>>(value);
		}

		public static List<WppSimpleModel> DeserializeToWppSimpleList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<WppSimpleModel>() : JsonConvert.DeserializeObject<List<WppSimpleModel>>(value);
		}

		public static MppModel DeserializeToMpp(this string value)
		{
			return string.IsNullOrEmpty(value) ? new MppModel() : JsonConvert.DeserializeObject<MppModel>(value);
		}

		public static List<MppModel> DeserializeToMppList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<MppModel>() : JsonConvert.DeserializeObject<List<MppModel>>(value);
		}

		public static MppChangesModel DeserializeToMppChanges(this string value)
		{
			return string.IsNullOrEmpty(value) ? new MppChangesModel() : JsonConvert.DeserializeObject<MppChangesModel>(value);
		}

		public static List<MppChangesModel> DeserializeToMppChangesList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<MppChangesModel>() : JsonConvert.DeserializeObject<List<MppChangesModel>>(value);
		}

		public static ReferenceModel DeserializeToReference(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ReferenceModel() : JsonConvert.DeserializeObject<ReferenceModel>(value);
		}

		public static List<ReferenceModel> DeserializeToReferenceList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ReferenceModel>() : JsonConvert.DeserializeObject<List<ReferenceModel>>(value);
		}

		public static ReferenceDetailModel DeserializeToRefDetail(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ReferenceDetailModel() : JsonConvert.DeserializeObject<ReferenceDetailModel>(value);
		}

		public static List<ReferenceDetailModel> DeserializeToRefDetailList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ReferenceDetailModel>() : JsonConvert.DeserializeObject<List<ReferenceDetailModel>>(value);
		}

		public static UserLogModel DeserializeToUserLog(this string value)
		{
			return string.IsNullOrEmpty(value) ? new UserLogModel() : JsonConvert.DeserializeObject<UserLogModel>(value);
		}

		public static List<UserLogModel> DeserializeToUserLogList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<UserLogModel>() : JsonConvert.DeserializeObject<List<UserLogModel>>(value);
		}

		public static ProductionCenterModel DeserializeToProductionCenter(this string value, int index, long locationID, long parentID)
		{
			if (string.IsNullOrEmpty(value))
			{
				return new ProductionCenterModel();
			}
			else
			{
				var result = JsonConvert.DeserializeObject<ProductionCenterModel>(value);
				result.Index = index;
				result.LocationID = locationID;
				result.ParentID = parentID;

				return result;
			}
		}

		public static DepartmentModel DeserializeToDepartment(this string value, int index, long locationID, long parentID)
		{
			if (string.IsNullOrEmpty(value))
			{
				return new DepartmentModel();
			}
			else
			{
				var result = JsonConvert.DeserializeObject<DepartmentModel>(value);
				result.Index = index;
				result.LocationID = locationID;
				result.ParentID = parentID;

				return result;
			}
		}

		public static SubDepartmentModel DeserializeToSubDepartment(this string value, int index, long locationID, long parentID)
		{
			if (string.IsNullOrEmpty(value))
			{
				return new SubDepartmentModel();
			}
			else
			{
				var result = JsonConvert.DeserializeObject<SubDepartmentModel>(value);
				result.Index = index;
				result.LocationID = locationID;
				result.ParentID = parentID;

				return result;
			}
		}

		public static List<WppChangesModel> DeserializeToWppChangesList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<WppChangesModel>() : JsonConvert.DeserializeObject<List<WppChangesModel>>(value);
		}

		public static WeeksModel DeserializeToWeek(this string value)
		{
			return string.IsNullOrEmpty(value) ? new WeeksModel() : JsonConvert.DeserializeObject<WeeksModel>(value);
		}

		public static List<WeeksModel> DeserializeToWeeksList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<WeeksModel>() : JsonConvert.DeserializeObject<List<WeeksModel>>(value);
		}

		public static LPHModel DeserializeToLPH(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LPHModel() : JsonConvert.DeserializeObject<LPHModel>(value);
		}

		public static List<LPHModel> DeserializeToLPHList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHModel>() : JsonConvert.DeserializeObject<List<LPHModel>>(value);
		}

		public static LPHLocationsModel DeserializeToLPHLocation(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LPHLocationsModel() : JsonConvert.DeserializeObject<LPHLocationsModel>(value);
		}

		public static List<LPHLocationsModel> DeserializeToLPHLocationList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHLocationsModel>() : JsonConvert.DeserializeObject<List<LPHLocationsModel>>(value);
		}

		public static LPHApprovalsModel DeserializeToLPHApproval(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LPHApprovalsModel() : JsonConvert.DeserializeObject<LPHApprovalsModel>(value);
		}

		public static List<LPHApprovalsModel> DeserializeToLPHApprovalList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHApprovalsModel>() : JsonConvert.DeserializeObject<List<LPHApprovalsModel>>(value);
		}

		public static LPHSubmissionsModel DeserializeToLPHSubmissions(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LPHSubmissionsModel() : JsonConvert.DeserializeObject<LPHSubmissionsModel>(value);
		}

		public static List<LPHSubmissionsModel> DeserializeToLPHSubmissionsList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHSubmissionsModel>() : JsonConvert.DeserializeObject<List<LPHSubmissionsModel>>(value);
		}

		public static List<LPHExtrasModel> DeserializeToLPHExtraList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHExtrasModel>() : JsonConvert.DeserializeObject<List<LPHExtrasModel>>(value);
		}

		public static LPHComponentsModel DeserializeToLPHComponent(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LPHComponentsModel() : JsonConvert.DeserializeObject<LPHComponentsModel>(value);
		}
		public static List<LPHComponentsModel> DeserializeToLPHComponentList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHComponentsModel>() : JsonConvert.DeserializeObject<List<LPHComponentsModel>>(value);
		}

		public static LPHValuesModel DeserializeToLPHValue(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LPHValuesModel() : JsonConvert.DeserializeObject<LPHValuesModel>(value);
		}

		public static List<LPHValuesModel> DeserializeToLPHValueList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHValuesModel>() : JsonConvert.DeserializeObject<List<LPHValuesModel>>(value);
		}

		public static LPHValueHistoriesModel DeserializeToLPHValueHistory(this string value)
		{
			return string.IsNullOrEmpty(value) ? new LPHValueHistoriesModel() : JsonConvert.DeserializeObject<LPHValueHistoriesModel>(value);
		}

		public static List<LPHValueHistoriesModel> DeserializeToLPHValueHistoryList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<LPHValueHistoriesModel>() : JsonConvert.DeserializeObject<List<LPHValueHistoriesModel>>(value);
		}

		public static PPLPHModel DeserializeToPPLPH(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHModel() : JsonConvert.DeserializeObject<PPLPHModel>(value);
		}
		public static List<PPLPHModel> DeserializeToPPLPHList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHModel>() : JsonConvert.DeserializeObject<List<PPLPHModel>>(value);
		}

		public static PPLPHLocationsModel DeserializeToPPLPHLocations(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHLocationsModel() : JsonConvert.DeserializeObject<PPLPHLocationsModel>(value);
		}
		public static List<PPLPHLocationsModel> DeserializeToPPLPHLocationsList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHLocationsModel>() : JsonConvert.DeserializeObject<List<PPLPHLocationsModel>>(value);
		}

		public static PPLPHSubmissionsModel DeserializeToPPLPHSubmissions(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHSubmissionsModel() : JsonConvert.DeserializeObject<PPLPHSubmissionsModel>(value);
		}
		public static List<PPLPHSubmissionsModel> DeserializeToPPLPHSubmissionsList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHSubmissionsModel>() : JsonConvert.DeserializeObject<List<PPLPHSubmissionsModel>>(value);
		}

		public static PPLPHApprovalsModel DeserializeToPPLPHApproval(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHApprovalsModel() : JsonConvert.DeserializeObject<PPLPHApprovalsModel>(value);
		}
		public static List<PPLPHApprovalsModel> DeserializeToPPLPHApprovalList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHApprovalsModel>() : JsonConvert.DeserializeObject<List<PPLPHApprovalsModel>>(value);
		}

		public static PPLPHComponentsModel DeserializeToPPLPHComponent(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHComponentsModel() : JsonConvert.DeserializeObject<PPLPHComponentsModel>(value);
		}
		public static List<PPLPHComponentsModel> DeserializeToPPLPHComponentList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHComponentsModel>() : JsonConvert.DeserializeObject<List<PPLPHComponentsModel>>(value);
		}

		public static PPLPHValuesModel DeserializeToPPLPHValue(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHValuesModel() : JsonConvert.DeserializeObject<PPLPHValuesModel>(value);
		}
		public static List<PPLPHValuesModel> DeserializeToPPLPHValueList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHValuesModel>() : JsonConvert.DeserializeObject<List<PPLPHValuesModel>>(value);
		}

		public static PPLPHValueHistoriesModel DeserializeToPPLPHValueHistories(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHValueHistoriesModel() : JsonConvert.DeserializeObject<PPLPHValueHistoriesModel>(value);
		}
		public static List<PPLPHValueHistoriesModel> DeserializeToPPLPHValueHistoriesList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHValueHistoriesModel>() : JsonConvert.DeserializeObject<List<PPLPHValueHistoriesModel>>(value);
		}

		public static PPLPHExtrasModel DeserializeToPPLPHExtras(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHExtrasModel() : JsonConvert.DeserializeObject<PPLPHExtrasModel>(value);
		}
		public static List<PPLPHExtrasModel> DeserializeToPPLPHExtrasList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHExtrasModel>() : JsonConvert.DeserializeObject<List<PPLPHExtrasModel>>(value);
		}

		public static TrainingModel DeserializeToTraining(this string value)
		{
			return string.IsNullOrEmpty(value) ? new TrainingModel() : JsonConvert.DeserializeObject<TrainingModel>(value);
		}
		public static List<TrainingModel> DeserializeToTrainingList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<TrainingModel>() : JsonConvert.DeserializeObject<List<TrainingModel>>(value);
		}
		public static PPLPHValueHistoriesModel DeserializeToPPLPHValueHistory(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPLPHValueHistoriesModel() : JsonConvert.DeserializeObject<PPLPHValueHistoriesModel>(value);
		}

		public static List<PPLPHValueHistoriesModel> DeserializeToPPLPHValueHistoryList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPLPHValueHistoriesModel>() : JsonConvert.DeserializeObject<List<PPLPHValueHistoriesModel>>(value);
		}

		public static ChecklistApprovalModel DeserializeToChecklistApproval(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistApprovalModel() : JsonConvert.DeserializeObject<ChecklistApprovalModel>(value);
		}

		public static List<ChecklistApprovalModel> DeserializeToChecklistApprovalList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistApprovalModel>() : JsonConvert.DeserializeObject<List<ChecklistApprovalModel>>(value);
		}

		public static ChecklistApproverModel DeserializeToChecklistApprover(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistApproverModel() : JsonConvert.DeserializeObject<ChecklistApproverModel>(value);
		}

		public static List<ChecklistApproverModel> DeserializeToChecklistApproverList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistApproverModel>() : JsonConvert.DeserializeObject<List<ChecklistApproverModel>>(value);
		}

		public static ChecklistComponentModel DeserializeToChecklistComponent(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistComponentModel() : JsonConvert.DeserializeObject<ChecklistComponentModel>(value);
		}

		public static List<ChecklistComponentModel> DeserializeToChecklistComponentList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistComponentModel>() : JsonConvert.DeserializeObject<List<ChecklistComponentModel>>(value);
		}

		public static ChecklistLocationModel DeserializeToChecklistLocation(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistLocationModel() : JsonConvert.DeserializeObject<ChecklistLocationModel>(value);
		}

		public static List<ChecklistLocationModel> DeserializeToChecklistLocationList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistLocationModel>() : JsonConvert.DeserializeObject<List<ChecklistLocationModel>>(value);
		}

		public static ChecklistModel DeserializeToChecklist(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistModel() : JsonConvert.DeserializeObject<ChecklistModel>(value);
		}

		public static List<ChecklistModel> DeserializeToChecklistList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistModel>() : JsonConvert.DeserializeObject<List<ChecklistModel>>(value);
		}

		public static ChecklistSubmitModel DeserializeToChecklistSubmit(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistSubmitModel() : JsonConvert.DeserializeObject<ChecklistSubmitModel>(value);
		}

		public static List<ChecklistSubmitModel> DeserializeToChecklistSubmitList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistSubmitModel>() : JsonConvert.DeserializeObject<List<ChecklistSubmitModel>>(value);
		}

		public static ChecklistValueHistoryModel DeserializeToChecklistValueHistory(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistValueHistoryModel() : JsonConvert.DeserializeObject<ChecklistValueHistoryModel>(value);
		}

		public static List<ChecklistValueHistoryModel> DeserializeToChecklistValueHistoryList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistValueHistoryModel>() : JsonConvert.DeserializeObject<List<ChecklistValueHistoryModel>>(value);
		}

		public static ChecklistValueModel DeserializeToChecklistValue(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ChecklistValueModel() : JsonConvert.DeserializeObject<ChecklistValueModel>(value);
		}

		public static List<ChecklistValueModel> DeserializeToChecklistValueList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ChecklistValueModel>() : JsonConvert.DeserializeObject<List<ChecklistValueModel>>(value);
		}

		public static ReportRemarksModel DeserializeToReportRemarks(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ReportRemarksModel() : JsonConvert.DeserializeObject<ReportRemarksModel>(value);
		}
		public static List<ReportRemarksModel> DeserializeToReportRemarksList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ReportRemarksModel>() : JsonConvert.DeserializeObject<List<ReportRemarksModel>>(value);
		}
		public static InputDailyModel DeserializeToInputDaily(this string value)
		{
			return string.IsNullOrEmpty(value) ? new InputDailyModel() : JsonConvert.DeserializeObject<InputDailyModel>(value);
		}
		public static List<InputDailyModel> DeserializeToInputDailyList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<InputDailyModel>() : JsonConvert.DeserializeObject<List<InputDailyModel>>(value);
		}

		public static InputOVModel DeserializeToInputOV(this string value)
		{
			return string.IsNullOrEmpty(value) ? new InputOVModel() : JsonConvert.DeserializeObject<InputOVModel>(value);
		}
		public static List<InputOVModel> DeserializeToInputOVList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<InputOVModel>() : JsonConvert.DeserializeObject<List<InputOVModel>>(value);
		}
		public static InputTargetModel DeserializeToInputTarget(this string value)
		{
			return string.IsNullOrEmpty(value) ? new InputTargetModel() : JsonConvert.DeserializeObject<InputTargetModel>(value);
		}
		public static List<InputTargetModel> DeserializeToInputTargetList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<InputTargetModel>() : JsonConvert.DeserializeObject<List<InputTargetModel>>(value);
		}
		public static PPReportYieldModel DeserializeToPPReportYield(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPReportYieldModel() : JsonConvert.DeserializeObject<PPReportYieldModel>(value);
		}
		public static PPReportYieldOvModel DeserializeToPPReportYieldOv(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPReportYieldOvModel() : JsonConvert.DeserializeObject<PPReportYieldOvModel>(value);
		}
		public static T DeserializeJson<T>(this string value)
		{
			return string.IsNullOrEmpty(value) ? default : JsonConvert.DeserializeObject<T>(value);
		}
		public static List<PPReportYieldModel> DeserializeToPPReportYieldList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPReportYieldModel>() : JsonConvert.DeserializeObject<List<PPReportYieldModel>>(value);
		}
		public static PPReportYieldWhiteModel DeserializeToPPReportYieldWhite(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPReportYieldWhiteModel() : JsonConvert.DeserializeObject<PPReportYieldWhiteModel>(value);
		}
		public static List<PPReportYieldWhiteModel> DeserializeToPPReportYieldWhiteList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPReportYieldWhiteModel>() : JsonConvert.DeserializeObject<List<PPReportYieldWhiteModel>>(value);
		}
		public static PPReportYieldKretekModel DeserializeToPPReportYieldKretek(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPReportYieldKretekModel() : JsonConvert.DeserializeObject<PPReportYieldKretekModel>(value);
		}
		public static List<PPReportYieldKretekModel> DeserializeToPPReportYieldKretekList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPReportYieldKretekModel>() : JsonConvert.DeserializeObject<List<PPReportYieldKretekModel>>(value);
		}
		public static List<PPReportYieldMCDietModel> DeserializeToPPReportYieldMCDietList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPReportYieldMCDietModel>() : JsonConvert.DeserializeObject<List<PPReportYieldMCDietModel>>(value);
		}

		public static PPReportDowntimeModel DeserializeToPPReportDowntimeModel(this string value)
		{
			return string.IsNullOrEmpty(value) ? new PPReportDowntimeModel() : JsonConvert.DeserializeObject<PPReportDowntimeModel>(value);
		}
		public static List<PPReportDowntimeModel> DeserializeToPPReportDowntimeModelList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPReportDowntimeModel>() : JsonConvert.DeserializeObject<List<PPReportDowntimeModel>>(value);
		}

		public static ReportKPICRRModel DeserializeToReportKPICRRModel(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ReportKPICRRModel() : JsonConvert.DeserializeObject<ReportKPICRRModel>(value);
		}
		public static List<ReportKPICRRModel> DeserializeToReportKPICRRModelList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ReportKPICRRModel>() : JsonConvert.DeserializeObject<List<ReportKPICRRModel>>(value);
		}

		public static ReportKPIDIMModel DeserializeToReportKPIDIMModel(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ReportKPIDIMModel() : JsonConvert.DeserializeObject<ReportKPIDIMModel>(value);
		}
		public static List<ReportKPIDIMModel> DeserializeToReportKPIDIMModelList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ReportKPIDIMModel>() : JsonConvert.DeserializeObject<List<ReportKPIDIMModel>>(value);
		}

		public static ReportKPIProdVolModel DeserializeToReportKPIProdVolModel(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ReportKPIProdVolModel() : JsonConvert.DeserializeObject<ReportKPIProdVolModel>(value);
		}
		public static List<ReportKPIProdVolModel> DeserializeToReportKPIProdVolModelList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ReportKPIProdVolModel>() : JsonConvert.DeserializeObject<List<ReportKPIProdVolModel>>(value);
		}

		public static ReportKPIYieldModel DeserializeToReportKPIYieldModel(this string value)
		{
			return string.IsNullOrEmpty(value) ? new ReportKPIYieldModel() : JsonConvert.DeserializeObject<ReportKPIYieldModel>(value);
		}
		public static List<ReportKPIYieldModel> DeserializeToReportKPIYieldModelList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<ReportKPIYieldModel>() : JsonConvert.DeserializeObject<List<ReportKPIYieldModel>>(value);
		}

        public static ReportKPIWorkHourModel DeserializeToReportKPIWorkHourModel(this string value)
        {
            return string.IsNullOrEmpty(value) ? new ReportKPIWorkHourModel() : JsonConvert.DeserializeObject<ReportKPIWorkHourModel>(value);
        }
        public static List<ReportKPIWorkHourModel> DeserializeToReportKPIWorkHourModelList(this string value)
        {
            return string.IsNullOrEmpty(value) ? new List<ReportKPIWorkHourModel>() : JsonConvert.DeserializeObject<List<ReportKPIWorkHourModel>>(value);
        }
        public static ReportKPIStickPerPackModel DeserializeToReportKPIStickPerPackModel(this string value)
        {
            return string.IsNullOrEmpty(value) ? new ReportKPIStickPerPackModel() : JsonConvert.DeserializeObject<ReportKPIStickPerPackModel>(value);
        }
        public static List<ReportKPIStickPerPackModel> DeserializeToReportKPIStickPerPackModelList(this string value)
        {
            return string.IsNullOrEmpty(value) ? new List<ReportKPIStickPerPackModel>() : JsonConvert.DeserializeObject<List<ReportKPIStickPerPackModel>>(value);
        }
        public static ReportKPIDustModel DeserializeToReportKPIDustModel(this string value)
        {
            return string.IsNullOrEmpty(value) ? new ReportKPIDustModel() : JsonConvert.DeserializeObject<ReportKPIDustModel>(value);
        }
        public static List<ReportKPIDustModel> DeserializeToReportKPIDustModelList(this string value)
        {
            return string.IsNullOrEmpty(value) ? new List<ReportKPIDustModel>() : JsonConvert.DeserializeObject<List<ReportKPIDustModel>>(value);
        }
        public static ReportKPITobaccoWeightModel DeserializeToReportKPITobaccoWeightModel(this string value)
        {
            return string.IsNullOrEmpty(value) ? new ReportKPITobaccoWeightModel() : JsonConvert.DeserializeObject<ReportKPITobaccoWeightModel>(value);
        }
        public static List<ReportKPITobaccoWeightModel> DeserializeToReportKPITobaccoWeightModelList(this string value)
        {
            return string.IsNullOrEmpty(value) ? new List<ReportKPITobaccoWeightModel>() : JsonConvert.DeserializeObject<List<ReportKPITobaccoWeightModel>>(value);
        }
        public static ReportKPIRipperInfoModel DeserializeToReportKPIRipperInfoModel(this string value)
        {
            return string.IsNullOrEmpty(value) ? new ReportKPIRipperInfoModel() : JsonConvert.DeserializeObject<ReportKPIRipperInfoModel>(value);
        }
        public static List<ReportKPIRipperInfoModel> DeserializeToReportKPIRipperInfoModelList(this string value)
        {
            return string.IsNullOrEmpty(value) ? new List<ReportKPIRipperInfoModel>() : JsonConvert.DeserializeObject<List<ReportKPIRipperInfoModel>>(value);
        }
        public static ReportKPICRRConversionModel DeserializeToReportKPICRRConversionModel(this string value)
        {
            return string.IsNullOrEmpty(value) ? new ReportKPICRRConversionModel() : JsonConvert.DeserializeObject<ReportKPICRRConversionModel>(value);
        }
        public static List<ReportKPICRRConversionModel> DeserializeToReportKPICRRConversionModelList(this string value)
        {
            return string.IsNullOrEmpty(value) ? new List<ReportKPICRRConversionModel>() : JsonConvert.DeserializeObject<List<ReportKPICRRConversionModel>>(value);
        }
        public static ReportKPITargetModel DeserializeToReportKPITargetModel(this string value)
        {
            return string.IsNullOrEmpty(value) ? new ReportKPITargetModel() : JsonConvert.DeserializeObject<ReportKPITargetModel>(value);
        }
        public static List<ReportKPITargetModel> DeserializeToReportKPITargetModelList(this string value)
        {
            return string.IsNullOrEmpty(value) ? new List<ReportKPITargetModel>() : JsonConvert.DeserializeObject<List<ReportKPITargetModel>>(value);
        }


        public static List<PPReportYieldKretekWestModel> DeserializeToPPReportYieldKretekWestList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPReportYieldKretekWestModel>() : JsonConvert.DeserializeObject<List<PPReportYieldKretekWestModel>>(value);
		}
		public static List<PPReportYieldWhiteWestModel> DeserializeToPPReportYieldWhiteWestList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PPReportYieldWhiteWestModel>() : JsonConvert.DeserializeObject<List<PPReportYieldWhiteWestModel>>(value);
		}
		public static List<OtherPemakaianBandrolModel> DeserializeToOtherPemakaianBandrolList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<OtherPemakaianBandrolModel>() : JsonConvert.DeserializeObject<List<OtherPemakaianBandrolModel>>(value);
		}
		public static List<PMIDPemakaianBandrolModel> DeserializeToPMIDPemakaianBandrolList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PMIDPemakaianBandrolModel>() : JsonConvert.DeserializeObject<List<PMIDPemakaianBandrolModel>>(value);
		}
		public static List<PMIDPemakaianBandrolSlofModel> DeserializeToPMIDPemakaianBandrolSlofList(this string value)
		{
			return string.IsNullOrEmpty(value) ? new List<PMIDPemakaianBandrolSlofModel>() : JsonConvert.DeserializeObject<List<PMIDPemakaianBandrolSlofModel>>(value);
		}
	}
}