using Fast.Domain.Entities;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration;

namespace Fast.Infra.Data
{
	public class FastAppContext : DbContext
	{
		public FastAppContext() : base("FastAppConn")
		{
			Configuration.LazyLoadingEnabled = false;
		}

		protected override void OnModelCreating(DbModelBuilder modelBuilder)
		{
			modelBuilder.Configurations.Add(new BrandConversionConfiguration());
			modelBuilder.Configurations.Add(new InputDailyConversionConfiguration());
		}

		public DbSet<Checklist> Checklists { get; set; }
		public DbSet<ChecklistApproval> ChecklistApprovals { get; set; }
		public DbSet<ChecklistApprover> ChecklistApprovers { get; set; }
		public DbSet<ChecklistSubmit> ChecklistSubmits { get; set; }
		public DbSet<ChecklistLocation> ChecklistsLocations { get; set; }
		public DbSet<ChecklistComponent> ChecklistsComponents { get; set; }
		public DbSet<ChecklistValue> ChecklistsValues { get; set; }
		public DbSet<ChecklistValueHistory> ChecklistsValueHistories { get; set; }
		public DbSet<JobTitle> JobTitles { get; set; }
		public DbSet<User> Users { get; set; }
		public DbSet<Role> Roles { get; set; }
		public DbSet<Location> Locations { get; set; }
		public DbSet<Wpp> WPPs { get; set; }
		public DbSet<WppStp> WPPSTPs { get; set; }
		public DbSet<Mpp> MPPs { get; set; }
		public DbSet<UserLog> UserLogs { get; set; }
		public DbSet<Menu> Menus { get; set; }
		public DbSet<AccessRight> AccessRights { get; set; }
		public DbSet<Skill> Skills { get; set; }
		public DbSet<Machine> Machines { get; set; }
		public DbSet<UserMachineType> UserMachineTypes { get; set; }
		public DbSet<UserMachine> UserMachines { get; set; }
		public DbSet<Reference> References { get; set; }
		public DbSet<ReferenceDetail> ReferenceDetails { get; set; }
		public DbSet<Calendar> Calendars { get; set; }
		public DbSet<CalendarHoliday> CalendaryHolidays { get; set; }
		public DbSet<EmployeeProfile> EmployeeProfiles { get; set; }
		public DbSet<EmployeeProfileAll> EmployeeProfileAlls { get; set; }
		public DbSet<EmployeeLeave> EmployeeLeaves { get; set; }
		public DbSet<EmployeeOvertime> EmployeeOvertimes { get; set; }
		public DbSet<MppChanges> MppChanges { get; set; }
		public DbSet<WppChanges> WppChanges { get; set; }
		public DbSet<Weeks> Weeks { get; set; }
		public DbSet<LocationMachineType> LocationMachineTypes { get; set; }
		public DbSet<LPH> LPHs { get; set; }
		public DbSet<LPHComponents> LPHComponents { get; set; }
		public DbSet<LPHLocations> LPHLocations { get; set; }
		public DbSet<LPHValues> LPHValues { get; set; }
		public DbSet<LPHValueHistories> LPHValueHistories { get; set; }
		public DbSet<LPHExtras> LPHExtras { get; set; }
		public DbSet<LPHApprovals> LPHApprovals { get; set; }
		public DbSet<LPHSubmissions> LPHSubmissions { get; set; }
		public DbSet<ManPower> ManPowers { get; set; }
		public DbSet<PPLPH> PPLPHs { get; set; }
		public DbSet<PPLPHComponents> PPLPHComponents { get; set; }
		public DbSet<PPLPHLocations> PPLPHLocations { get; set; }
		public DbSet<PPLPHValues> PPLPHValues { get; set; }
		public DbSet<PPLPHValueHistories> PPLPHValueHistories { get; set; }
		public DbSet<PPLPHExtras> PPLPHExtras { get; set; }
		public DbSet<PPLPHApprovals> PPLPHApprovals { get; set; }
		public DbSet<PPLPHSubmissions> PPLPHSubmissions { get; set; }
		public DbSet<Training> Trainings { get; set; }
		public DbSet<Brand> Brands { get; set; }
		public DbSet<Blend> Blends { get; set; }
		public DbSet<UserRole> UserRoles { get; set; }
		public DbSet<MaterialCode> MaterialCodes { get; set; }
		public DbSet<WppPrim> WppPrims { get; set; }
		public DbSet<BrandConversion> BrandConversions { get; set; }
		public DbSet<ReportRemarks> ReportRemarks { get; set; }
		public DbSet<TrainingTitle> TrainingTitles { get; set; }
		public DbSet<TrainingMachineType> TrainingMachineTypes { get; set; }
		public DbSet<InputDaily> InputDailies { get; set; }
		public DbSet<MealRequest> MealRequests { get; set; }
		public DbSet<ShuttleRequest> ShuttleRequests { get; set; }
		public DbSet<InputOV> InputOVs { get; set; }
		public DbSet<MachineAllocation> MachineAllocations { get; set; }
		public DbSet<WppPrimary> WppPrimaries { get; set; }
		public DbSet<InputTarget> InputTargets { get; set; }
		public DbSet<PPReportYieldOvs> PPReportYieldOvs { get; set; }
		public DbSet<PPReportYieldTargets> PPReportYieldTargets { get; set; }
		public DbSet<PPReportYieldIMLs> PPReportYieldIMLs { get; set; }
		public DbSet<PPReportYields> PPReportYields { get; set; }
		public DbSet<PPReportYieldMCDiets> PPReportYieldMCDiets { get; set; }
		public DbSet<PPReportYieldKreteks> PPReportYieldKreteks { get; set; }
		public DbSet<PPReportYieldWhites> PPReportYieldWhites { get; set; }
		public DbSet<ReportKPICRR> ReportKPICRRs { get; set; }
		public DbSet<ReportKPIDIM> ReportKPIDIMs { get; set; }
		public DbSet<ReportKPIProdVol> ReportKPIProdVols { get; set; }
		public DbSet<ReportKPIYield> ReportKPIYields { get; set; }
		public DbSet<ReportKPICRRConversion> ReportKPICRRConversions { get; set; }
		public DbSet<ReportKPIWorkHour> ReportKPIWorkHours { get; set; }
		public DbSet<ReportKPIDust> ReportKPIDusts { get; set; }
		public DbSet<ReportKPITarget> ReportKPITargets { get; set; }
		public DbSet<ReportKPIStickPerPack> ReportKPIStickPerPacks { get; set; }
		public DbSet<ReportKPITobaccoWeight> ReportKPITobaccoWeights { get; set; }
		public DbSet<ReportKPIRipperInfo> ReportKPIRipperInfos { get; set; }
	}
}

public class InputDailyConversionConfiguration : EntityTypeConfiguration<InputDaily>
{
	public InputDailyConversionConfiguration()
	{
		ToTable("InputDailies");
		HasKey(x => x.ID).Property(x => x.ID).HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

		Property(x => x.MTBFValue).HasPrecision(precision: 18, scale: 8);
		Property(x => x.CPQIValue).HasPrecision(precision: 18, scale: 8);
		Property(x => x.VQIValue).HasPrecision(precision: 18, scale: 8);
		Property(x => x.WorkingValue).HasPrecision(precision: 18, scale: 8);
		Property(x => x.UptimeValue).HasPrecision(precision: 18, scale: 8);
		Property(x => x.STRSValue).HasPrecision(precision: 18, scale: 8);
		Property(x => x.ProdVolumeValue).HasPrecision(precision: 18, scale: 8);
		Property(x => x.CRRValue).HasPrecision(precision: 18, scale: 8);
	}
}

public class BrandConversionConfiguration : EntityTypeConfiguration<BrandConversion>
{
	public BrandConversionConfiguration()
	{
		ToTable("BrandConversions");
		HasKey(x => x.ID).Property(x => x.ID).HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

		Property(x => x.Value1).HasPrecision(precision: 15, scale: 4);
		Property(x => x.Value2).HasPrecision(precision: 15, scale: 4);
	}
}