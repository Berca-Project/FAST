using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.Web.Models;
using Fast.Web.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Fast.Web.Utils
{
    public static class DropDownHelper
    {
        public static List<SelectListItem> BuildEmpty()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "",
                Value = ""
            });

            return _menuList;
        }

        public static MultiSelectList BuildMultiEmpty()
        {
            MultiSelectList multilist = new MultiSelectList(new List<SelectListItem>(), "ID", "Code");
            return multilist;
        }

        public static List<SelectListItem> BuildEmptyList()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "- Please Select -",
                Value = "0"
            });

            return _menuList;
        }

        public static List<SelectListItem> BuildEmptyListDepartment()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "- Please Select Prod. Center First -",
                Value = "0"
            });

            return _menuList;
        }

        public static List<SelectListItem> BuildEmptyListSubDepartment()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "- Please Select Department First -",
                Value = "0"
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownStatusMpp()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "Normal",
                Value = "Normal"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "LongShift1",
                Value = "LongShift1"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "LongShift2",
                Value = "LongShift2"
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownSimpleStatusMpp()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "Normal",
                Value = "Normal"
            });
            _menuList.Add(new SelectListItem
            {
                Text = "LongShift",
                Value = "LongShift"
            });

            return _menuList;
        }

        public static string ExtractLocation(ILocationAppService _locationAppService, long? locationID)
        {
            long countryID = 0;
            long pcID = 0;
            long depID = 0;
            long subDepID = 0;
            return ExtractLocation(_locationAppService, locationID, out countryID, out pcID, out depID, out subDepID);
        }

        public static string ExtractLocation(
            ILocationAppService _locationAppService,
            long? locationID,
            out long countryID,
            out long pcID,
            out long depID,
            out long subDepID)
        {

            countryID = 0;
            pcID = 0;
            depID = 0;
            subDepID = 0;

            string completeLocation = string.Empty;

            if (locationID.HasValue)
            {
                string loc = _locationAppService.GetById(locationID.Value, true);
                LocationModel locModel = loc.DeserializeToLocation();
                if (!string.IsNullOrEmpty(locModel.Code))
                    completeLocation = locModel.Code;

                if (locModel.ParentID > 0)
                {
                    loc = _locationAppService.GetById(locModel.ParentID, true);
                    locModel = loc.DeserializeToLocation();
                    if (!string.IsNullOrEmpty(locModel.Code))
                        completeLocation = locModel.Code + "-" + completeLocation;

                    if (locModel.ParentID > 0)
                    {
                        loc = _locationAppService.GetById(locModel.ParentID, true);
                        locModel = loc.DeserializeToLocation();
                        if (!string.IsNullOrEmpty(locModel.Code))
                            completeLocation = locModel.Code + "-" + completeLocation;

                        if (locModel.ParentID > 0)
                        {
                            loc = _locationAppService.GetById(locModel.ParentID, true);
                            locModel = loc.DeserializeToLocation();
                            if (!string.IsNullOrEmpty(locModel.Code))
                                completeLocation = locModel.Code + "-" + completeLocation;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(completeLocation))
                {
                    int index = 1;
                    long tempDepID = 0;
                    long tempPCID = 0;
                    string[] temp = completeLocation.Split('-');
                    foreach (var code in temp)
                    {
                        if (index == 1)
                        {
                            string locs = _locationAppService.FindByNoTracking("Code", code, true);
                            List<LocationModel> locModelList = locs.DeserializeToLocationList();
                            if (locModelList.Count > 0)
                            {
                                locModel = locModelList.Where(x => x.ParentCode == null).FirstOrDefault();
                                if (locModel != null)
                                    countryID = locModel.ID;
                            }
                        }
                        else if (index == 2)
                        {
                            string locs = _locationAppService.FindByNoTracking("Code", code, true);
                            List<LocationModel> locModelList = locs.DeserializeToLocationList();
                            if (locModelList.Count > 0)
                            {
                                locModel = locModelList.Where(x => x.ParentCode == temp[0]).FirstOrDefault();
                                if (locModel != null)
                                    pcID = tempPCID = locModel.ID;
                            }
                        }
                        else if (index == 3)
                        {
                            string locs = _locationAppService.FindByNoTracking("Code", code, true);
                            List<LocationModel> locModelList = locs.DeserializeToLocationList();
                            if (locModelList.Count > 0)
                            {
                                locModel = locModelList.Where(x => x.ParentID == tempPCID).FirstOrDefault();
                                if (locModel != null)
                                    depID = tempDepID = locModel.ID;
                            }
                        }
                        else if (index == 4)
                        {
                            string locs = _locationAppService.FindByNoTracking("Code", code, true);
                            List<LocationModel> locModelList = locs.DeserializeToLocationList();
                            if (locModelList.Count > 0)
                            {
                                locModel = locModelList.Where(x => x.ParentID == tempDepID).FirstOrDefault();
                                if (locModel != null)
                                    subDepID = locModel.ID;
                            }
                        }

                        index++;
                    }
                }
            }

            return completeLocation;
        }

        public static List<SelectListItem> GetSubDepartmentByDepartmentID(
            ILocationAppService _locationAppService,
            IReferenceAppService _referenceAppService,
            long id,
            string selectedId = null)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string pcs = _locationAppService.FindByNoTracking("ParentID", id.ToString(), true);
            List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.SubDep).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refPC.DeserializeToRefDetailList();

            _menuList.Add(new SelectListItem
            {
                Text = "- Please Select -",
                Value = "0"
            });

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    if (item.ID.ToString() == selectedId)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString(),
                            Selected = true
                        });
                    }
                    else
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString()
                        });
                    }
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> GetDepartmentByProdCenterID(
            ILocationAppService _locationAppService,
            IReferenceAppService _referenceAppService,
            long id,
            string selectedId = null)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string pcs = _locationAppService.FindByNoTracking("ParentID", id.ToString(), true);
            List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.Dep).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refPC.DeserializeToRefDetailList();

            _menuList.Add(new SelectListItem
            {
                Text = "- Please Select -",
                Value = "0"
            });

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    if (item.ID.ToString() == selectedId)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString(),
                            Selected = true
                        });
                    }
                    else
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString()
                        });
                    }
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> GetMachineTypeByProductionCenterID(
            ILocationAppService _locationAppService,
            ILocationMachineTypeAppService _locationMachineTypeAppService,
            IReferenceAppService _referenceAppService,
            long prodCenterID)
        {
            string departments = _locationAppService.FindBy("ParentID", prodCenterID, true);
            List<LocationModel> depList = departments.DeserializeToLocationList();
            List<LocationMachineTypeModel> result = new List<LocationMachineTypeModel>();

            foreach (var item in depList)
            {
                string subDeps = _locationAppService.FindBy("ParentID", item.ID, true);
                List<LocationModel> subDepList = subDeps.DeserializeToLocationList();

                foreach (var sd in subDepList)
                {
                    string machineTypes = _locationMachineTypeAppService.FindBy("LocationID", sd.ID, true);
                    result.AddRange(machineTypes.DeserializeToLocationMachineTypeList());
                }
            }

            result = result.GroupBy(p => p.MachineTypeID).Select(g => g.First()).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();
            List<SelectListItem> _machineTypeList = BindDropDownMachineType(_referenceAppService);

            foreach (var item in _machineTypeList)
            {
                if (result.Any(x => x.MachineTypeID.ToString() == item.Value))
                {
                    item.Value = "#" + item.Value;
                    _menuList.Add(item);
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> GetMachineTypeByDepartmentID(
            ILocationAppService _locationAppService,
            ILocationMachineTypeAppService _locationMachineTypeAppService,
            IReferenceAppService _referenceAppService,
            long departmentID)
        {
            string depCode = _locationAppService.FindBy("ID", departmentID, true);
            List<LocationModel> sdList = depCode.DeserializeToLocationList();
            List<LocationMachineTypeModel> result = new List<LocationMachineTypeModel>();

            //string subDeps = _locationAppService.FindBy("ParentID", departmentID, true);
            //List<LocationModel> sdList = subDeps.DeserializeToLocationList();
            //List<LocationMachineTypeModel> result = new List<LocationMachineTypeModel>();

            //         foreach (var sd in sdList)
            //{
            //	string machineTypes = _locationMachineTypeAppService.FindBy("LocationID", sd.ID, true);
            //	result.AddRange(machineTypes.DeserializeToLocationMachineTypeList());
            //}

            //result = result.GroupBy(p => p.MachineTypeID).Select(g => g.First()).ToList();
            string kode = "";
            foreach (var sd in sdList)
            {
                kode = sd.Code;
                break;
            }

            List<SelectListItem> _menuList = new List<SelectListItem>();
            List<SelectListItem> _machineTypeList = BindDropDownMachineType(_referenceAppService, kode);

            foreach (var item in _machineTypeList)
            {
                //if (result.Any(x => x.MachineTypeID.ToString() == item.Value))
                //{
                item.Value = "#" + item.Value;
                _menuList.Add(item);
                //}
            }

            return _menuList;
        }


        public static List<SelectListItem> GetMachineTypeBySubDepartmentID(
            ILocationMachineTypeAppService _locationMachineTypeAppService,
            IReferenceAppService _referenceAppService,
            long subdepartmentID)
        {
            List<LocationMachineTypeModel> result = new List<LocationMachineTypeModel>();

            string machineTypes = _locationMachineTypeAppService.FindBy("LocationID", subdepartmentID, true);
            result.AddRange(machineTypes.DeserializeToLocationMachineTypeList());

            result = result.GroupBy(p => p.MachineTypeID).Select(g => g.First()).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();
            List<SelectListItem> _machineTypeList = BindDropDownMachineType(_referenceAppService);

            foreach (var item in _machineTypeList)
            {
                if (result.Any(x => x.MachineTypeID.ToString() == item.Value))
                {
                    item.Value = "#" + item.Value;
                    _menuList.Add(item);
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> GetProductionCenterByCountryID(
            ILocationAppService _locationAppService,
            IReferenceAppService _referenceAppService,
            long id,
            string selectedId = null)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string pcs = _locationAppService.FindByNoTracking("ParentID", id.ToString(), true);
            List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.ProdCenter).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refPC.DeserializeToRefDetailList();

            _menuList.Add(new SelectListItem
            {
                Text = "- Please Select -",
                Value = ""
            });

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    if (item.ID.ToString() == selectedId)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString(),
                            Selected = true
                        });
                    }
                    else
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString()
                        });
                    }
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> GetCountryList(ILocationAppService _locationAppService, IReferenceAppService _referenceAppService, string selectedId = null)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string countries = _locationAppService.FindByNoTracking("ParentID", "0", true);
            List<LocationModel> locModelList = countries.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refCountries = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.Country).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refCountries.DeserializeToRefDetailList();

            _menuList.Add(new SelectListItem
            {
                Text = "- Please Select -",
                Value = "0"
            });

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    if (item.ID.ToString() == selectedId)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString(),
                            Selected = true
                        });
                    }
                    else
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString()
                        });
                    }
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownMachineCode(IMachineAppService service)
        {
            string machines = service.GetAll(true);
            List<MachineModel> machineModelList = machines.DeserializeToMachineList();
            machineModelList = machineModelList.OrderBy(x => x.Code).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in machineModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownMachineMakerCode(IMachineAppService service)
        {
            string machines = service.GetAll(true);
            List<MachineModel> machineModelList = machines.DeserializeToMachineList();
            machineModelList = machineModelList.Where(x => x.SubProcess == "Maker").ToList();
            machineModelList = machineModelList.GroupBy(p => p.Code).Select(g => g.First()).OrderBy(x => x.Code).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in machineModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownMachine(IMachineAppService service)
        {
            string machines = service.GetAll(true);
            List<MachineModel> machineModelList = machines.DeserializeToMachineList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in machineModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = item.ID.ToString()
                });
            }

            return _menuList;
        }

        public static MultiSelectList BindDropDownMultiRole(IRoleAppService service)
        {
            string roles = service.GetAll(true);
            List<RoleModel> roleList = roles.DeserializeToRoleList();
            roleList = roleList.GroupBy(x => x.Name).Select(g => g.First()).OrderBy(x => x.Name).ToList();

            MultiSelectList list = new MultiSelectList(roleList, "Name", "Name");

            return list;
        }

        public static MultiSelectList BindDropDownMultiRole(IRoleAppService service, List<string> selectedRoles)
        {
            string roles = service.GetAll(true);
            List<RoleModel> roleList = roles.DeserializeToRoleList();
            roleList = roleList.GroupBy(x => x.Name).Select(g => g.First()).OrderBy(x => x.Name).ToList();

            MultiSelectList list = new MultiSelectList(roleList, "Name", "Name", selectedRoles);

            return list;
        }

        public static MultiSelectList BindDropDownMultiMachine(IMachineAppService service)
        {
            string machines = service.GetAll(true);
            List<MachineModel> machineModelList = machines.DeserializeToMachineList();
            machineModelList = machineModelList.GroupBy(x => x.Code).Select(g => g.First()).OrderBy(x => x.Code).ToList();

            MultiSelectList list = new MultiSelectList(machineModelList, "ID", "Code");

            return list;
        }

        public static MultiSelectList BindDropDownMultiMachine(IMachineAppService service, List<long> selectedIds)
        {
            string machines = service.GetAll(true);
            List<MachineModel> machineModelList = machines.DeserializeToMachineList();
            machineModelList = machineModelList.GroupBy(x => x.Code).Select(g => g.First()).OrderBy(x => x.Code).ToList();

            MultiSelectList list = new MultiSelectList(machineModelList, "ID", "Code", selectedIds);

            return list;
        }

        public static MultiSelectList BindDropDownMultiMachineType(IReferenceAppService service)
        {
            string type = service.GetDetailAll(ReferenceEnum.MachineType, true);
            List<ReferenceDetailModel> typeModelList = type.DeserializeToRefDetailList();
            typeModelList = typeModelList.OrderBy(x => x.Code).ToList();
            MultiSelectList list = new MultiSelectList(typeModelList, "ID", "Code");

            return list;
        }

        public static MultiSelectList BindDropDownMultiMachineType(IReferenceAppService service, List<long> selectedIds)
        {
            string type = service.GetDetailAll(ReferenceEnum.MachineType, true);
            List<ReferenceDetailModel> typeModelList = type.DeserializeToRefDetailList();
            typeModelList = typeModelList.OrderBy(x => x.Code).ToList();

            MultiSelectList list = new MultiSelectList(typeModelList, "ID", "Code", selectedIds.ToArray());

            return list;
        }

        public static List<SelectListItem> BindDropDownMachineType(IReferenceAppService service, string kode = "")
        {
            string type = service.GetDetailAll(ReferenceEnum.MachineType, true);
            List<ReferenceDetailModel> typeModelList = type.DeserializeToRefDetailList();
            if (kode != "")
                typeModelList = typeModelList.Where(x => x.Description == kode).ToList();

            typeModelList = typeModelList.OrderBy(x => x.Code).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in typeModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = item.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownCostCenter(IReferenceAppService service)
        {
            string type = service.GetDetailAll(ReferenceEnum.CostCenter, true);
            List<ReferenceDetailModel> typeModelList = type.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in typeModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code + " - " + item.Description,
                    Value = item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownShift(bool isIncludePleaseSelect = false)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            if (isIncludePleaseSelect)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = UIResources.PleaseSelect,
                    Value = "0"
                });
            }

            _menuList.Add(new SelectListItem
            {
                Text = "1",
                Value = "1"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "2",
                Value = "2"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "3",
                Value = "3"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "NS",
                Value = "NS"
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownGuestType()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "External",
                Value = "External"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Internal",
                Value = "Internal"
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownCanteen(IReferenceAppService service, bool byID = false)
        {
            string reference = service.GetBy("Name", "Canteen", true);
            ReferenceModel refModel = reference.DeserializeToReference();

            string canteens = service.FindDetailBy("ReferenceID", refModel.ID, true);
            List<ReferenceDetailModel> canteenModelList = canteens.DeserializeToRefDetailList().OrderBy(x => x.Description).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            if (!byID)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = UIResources.PleaseSelect,
                    Value = "0"
                });
            }

            foreach (var item in canteenModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = byID ? item.ID.ToString() : item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownHalte(IReferenceAppService service, bool byID = false)
        {
            string reference = service.GetBy("Name", "Halte", true);
            ReferenceModel refModel = reference.DeserializeToReference();

            string canteens = service.FindDetailBy("ReferenceID", refModel.ID, true);
            List<ReferenceDetailModel> canteenModelList = canteens.DeserializeToRefDetailList().OrderBy(x => x.Description).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            if (!byID)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = UIResources.PleaseSelect,
                    Value = "0"
                });
            }

            foreach (var item in canteenModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = byID ? item.ID.ToString() : item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownRoleAddNew(IRoleAppService service)
        {
            string roles = service.GetAll(true);
            List<RoleModel> roleList = roles.DeserializeToRoleList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "- Create New -",
                Value = "0",
            });

            bool isFirstValue = true;
            foreach (var role in roleList)
            {
                if (isFirstValue)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = role.Name,
                        Value = role.Name,
                        Selected = true
                    });

                    isFirstValue = false;
                }
                else
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = role.Name,
                        Value = role.Name
                    });
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownRole(IRoleAppService service)
        {
            string roles = service.GetAll(true);
            List<RoleModel> roleList = roles.DeserializeToRoleList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            bool isFirstValue = true;
            foreach (var role in roleList)
            {
                if (isFirstValue)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = role.Name,
                        Value = role.Name,
                        Selected = true
                    });

                    isFirstValue = false;
                }
                else
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = role.Name,
                        Value = role.Name
                    });
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownParentMenu(IMenuAppService service)
        {
            string menus = service.GetAll(true);
            List<MenuModel> menuList = menus.DeserializeToMenuList();
            List<SelectListItem> _menuList = new List<SelectListItem>();
            menuList = menuList.Where(x => x.IsParent).ToList();

            foreach (var menu in menuList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = menu.Name,
                    Value = menu.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownMenu(IMenuAppService service)
        {
            string menus = service.GetAll(true);
            List<MenuModel> menuList = menus.DeserializeToMenuList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var menu in menuList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = menu.Name,
                    Value = menu.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownGroupType(IReferenceDetailAppService service)
        {
            string refDetails = service.FindBy("ReferenceID", "1", true);
            List<ReferenceDetailModel> redDetailModels = refDetails.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var detail in redDetailModels)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = detail.Code,
                    Value = detail.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownProdCenterCode(IReferenceAppService refService, ILocationAppService locService)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            // get production center list
            string pcs = locService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> pcList = pcs.DeserializeToLocationList();
            foreach (var pc in pcList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = pc.ProdCenterFullCode,
                    Value = pc.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownDepartmentCode(IReferenceAppService refService, ILocationAppService locService)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            // get production center list
            string pcs = locService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> pcList = pcs.DeserializeToLocationList();
            foreach (var pc in pcList)
            {
                string departments = locService.FindBy("ParentID", pc.ID, true);
                List<LocationModel> departmentList = departments.DeserializeToLocationList();

                foreach (var department in departmentList)
                {
                    department.CountryCode = pc.ParentCode;

                    _menuList.Add(new SelectListItem
                    {
                        Text = department.DepartementFullCode,
                        Value = department.ID.ToString()
                    });
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownSubDepartmentCode(IReferenceAppService refService, ILocationAppService locService)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            // get production center list
            string pcs = locService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> pcList = pcs.DeserializeToLocationList();
            foreach (var pc in pcList)
            {
                string departments = locService.FindBy("ParentID", pc.ID, true);
                List<LocationModel> departmentList = departments.DeserializeToLocationList();

                foreach (var department in departmentList)
                {
                    string subDeps = locService.FindBy("ParentID", department.ID, true);
                    List<LocationModel> subDepList = subDeps.DeserializeToLocationList();

                    foreach (var sd in subDepList)
                    {
                        sd.CountryCode = pc.ParentCode;
                        sd.ProductionCenterCode = pc.Code;

                        _menuList.Add(new SelectListItem
                        {
                            Text = sd.SubDepartementFullCode,
                            Value = sd.ID.ToString()
                        });
                    }
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownLocationTypeLocation(IReferenceAppService service)
        {
            string locType = service.GetDetailAll(ReferenceEnum.LocationType, true);
            List<ReferenceDetailModel> locTypeModels = locType.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();
            locTypeModels = locTypeModels.OrderBy(x => x.Description).ToList();

            _menuList.Add(new SelectListItem
            {
                Text = UIResources.PleaseSelect,
                Value = "0"
            });

            foreach (var item in locTypeModels)
            {
                if (!item.Code.Equals("SubDep"))
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = item.Description,
                        Value = item.ID.ToString()
                    });
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownYear()
        {

            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = DateTime.Now.Year.ToString(),
                Value = DateTime.Now.Year.ToString()
            });

            _menuList.Add(new SelectListItem
            {
                Text = DateTime.Now.AddYears(1).Year.ToString(),
                Value = DateTime.Now.AddYears(1).Year.ToString()
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownYearWpp()
        {

            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = DateTime.Now.AddYears(-1).Year.ToString(),
                Value = DateTime.Now.AddYears(-1).Year.ToString()
            });

            _menuList.Add(new SelectListItem
            {
                Text = DateTime.Now.Year.ToString(),
                Value = DateTime.Now.Year.ToString()
            });

            _menuList.Add(new SelectListItem
            {
                Text = DateTime.Now.AddYears(1).Year.ToString(),
                Value = DateTime.Now.AddYears(1).Year.ToString()
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownCountry(IReferenceAppService service)
        {
            string countries = service.GetDetailAll(ReferenceEnum.Country, true);
            List<ReferenceDetailModel> countryModels = countries.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            //TODO
            foreach (var item in countryModels)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code + " - " + item.Description,
                    Value = "1"
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownMachineCategory(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string reference = service.GetBy("Name", "MachineCategory");
            ReferenceModel referenceModel = reference.DeserializeToReference();
            string gnames = service.FindDetailBy("ReferenceID", referenceModel.ID, true);
            List<ReferenceDetailModel> gnamesModel = gnames.DeserializeToRefDetailList();

            foreach (var item in gnamesModel)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownGroupName(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string reference = service.GetBy("Name", "GroupName");
            ReferenceModel referenceModel = reference.DeserializeToReference();
            string gnames = service.FindDetailBy("ReferenceID", referenceModel.ID, true);
            List<ReferenceDetailModel> gnamesModel = gnames.DeserializeToRefDetailList();

            foreach (var item in gnamesModel)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownCluster(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string reference = service.GetBy("Name", "Cluster");
            ReferenceModel referenceModel = reference.DeserializeToReference();
            string gnames = service.FindDetailBy("ReferenceID", referenceModel.ID, true);
            List<ReferenceDetailModel> gnamesModel = gnames.DeserializeToRefDetailList();

            foreach (var item in gnamesModel)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.Code,
                    Value = item.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownSearchBy()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "Employee ID",
                Value = "ID"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Employee Name",
                Value = "Name"
            });

            return _menuList;

        }

        public static List<SelectListItem> BindDropDownOvertimeCategory()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string[] listCategory = new string[] {
                "Maintenance Activity",
                "Daily Work",
                "Other",
                "Emergency / Force Major",
                "Backup Leave",
                "Training",
                "Project",
                "Rework & WIP Activity"
            };

            foreach (var item in listCategory)
            {
                _menuList.Add(new SelectListItem { Text = item, Value = item });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownTrainingTitle(ITrainingTitleAppService service)
        {
            string trainingTitleList = service.GetAll(true);
            List<TrainingTitleModel> trainingTitleModelList = trainingTitleList.DeserializeToTrainingTitleList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in trainingTitleModelList)
            {
                _menuList.Add(new SelectListItem { Text = item.Title, Value = item.ID.ToString() });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownUserStatus()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "Active",
                Value = "true"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Inactive",
                Value = "false"
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownUserAdmin()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "True",
                Value = "true"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "False",
                Value = "false"
            });

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownJobTitleAll(IJobTitleAppService service)
        {
            string jobTitle = service.FindBy("IsDeleted", "0", true);
            List<JobTitleModel> jobTitleList = jobTitle.DeserializeToJobTitleList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var role in jobTitleList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = role.Title,
                    Value = role.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownJobTitleUnAssigned(IJobTitleAppService service)
        {
            string jobTitle = service.FindBy("IsDeleted", "0", true);
            List<JobTitleModel> jobTitleList = jobTitle.DeserializeToJobTitleList();
            jobTitleList = jobTitleList.Where(x => x.RoleName == null).OrderBy(x => x.Title).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var role in jobTitleList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = role.Title,
                    Value = role.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownJobTitleByRole(IJobTitleAppService service, string roleName)
        {
            string jobTitle = service.FindByNoTracking("IsDeleted", "0", true);
            List<JobTitleModel> jobTitleList = jobTitle.DeserializeToJobTitleList();
            jobTitleList = jobTitleList.Where(x => x.RoleName == roleName).OrderBy(x => x.Title).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var jt in jobTitleList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = jt.Title,
                    Value = jt.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownJobTitle(IJobTitleAppService service, string roleName = null)
        {
            string jobTitle = service.FindBy("IsDeleted", "0", true);
            List<JobTitleModel> jobTitleList = jobTitle.DeserializeToJobTitleList();
            jobTitleList = jobTitleList.Where(x => !x.Title.Contains("OS -")).OrderBy(x => x.Title).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();
            if (!string.IsNullOrEmpty(roleName))
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "- Create New -",
                    Value = "0"
                });

                jobTitleList = jobTitleList.Where(x => x.RoleName == roleName).ToList();
            }

            foreach (var jt in jobTitleList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = jt.Title,
                    Value = jt.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownJobTitleOS(IJobTitleAppService service, bool isIncludeNew = true)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("Title", "OS -", Operator.Contains));
            filters.Add(new QueryFilter("IsDeleted", "0"));

            string jobTitle = service.Find(filters);
            List<JobTitleModel> jobTitleList = jobTitle.DeserializeToJobTitleList().OrderBy(x => x.Title).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();

            if (isIncludeNew)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "- Create New -",
                    Value = "0"
                });
            }

            foreach (var jt in jobTitleList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = jt.Title,
                    Value = jt.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownProdCenter(IReferenceAppService service)
        {
            string pcList = service.GetDetailAll(ReferenceEnum.ProdCenter, true);
            List<ReferenceDetailModel> pcModelList = pcList.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var pc in pcModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = pc.Description,
                    Value = pc.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownUser(IUserAppService service)
        {
            string user = service.GetAll(true);
            List<UserModel> userModelList = user.DeserializeToUserList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in userModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = item.UserName,
                    Value = item.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownUserEmp(IUserAppService service, IEmployeeAppService empService)
        {
            string user = service.GetAll(true);
            List<UserModel> userModelList = user.DeserializeToUserList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in userModelList)
            {
                string emp = empService.GetBy("EmployeeID", item.EmployeeID, true);
                EmployeeModel empModel = emp.DeserializeToEmployee();

                _menuList.Add(new SelectListItem
                {
                    Text = empModel.FullName,
                    Value = item.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownUserEmpPosition(IUserAppService service, IEmployeeAppService empService)
        {
            string user = service.GetAll(true);
            List<UserModel> userModelList = user.DeserializeToUserList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var item in userModelList)
            {
                string emp = empService.GetBy("EmployeeID", item.EmployeeID, true);
                EmployeeModel empModel = emp.DeserializeToEmployee();

                _menuList.Add(new SelectListItem
                {
                    Text = empModel.FullName + " - " + empModel.PositionDesc,
                    Value = item.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownLinkUp(IReferenceAppService service, IMachineAppService machineService)
        {
            string dataList = service.GetDetailAll(ReferenceEnum.LinkUp, true);
            List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "- Others -",
                Value = "0"
            });

            foreach (var data in dataModelList)
            {
                string machineList = machineService.FindBy("LinkUp", data.Code, true);
                List<MachineModel> machineModelList = machineList.DeserializeToMachineList();
                var res = String.Join(" - ", machineModelList.Select(x => x.Code));

                _menuList.Add(new SelectListItem
                {
                    Text = data.Code + " [ " + res + " ]",
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownEmployee(IEmployeeAppService service)
        {
            string dataList = service.GetAll(true);
            List<EmployeeModel> dataModelList = dataList.DeserializeToEmployeeList().OrderBy(x => x.FullName).ToList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.FullName + " ( " + data.PositionDesc + " )",
                    Value = data.EmployeeID.Trim()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownReference(IReferenceAppService service)
        {
            string referenceList = service.GetAll(true);
            List<ReferenceModel> references = referenceList.DeserializeToReferenceList();
            references = references.Where(
                x => !x.Name.Equals("Country") &&
                !x.Name.Equals("PC") &&
                !x.Name.Equals("Dep") &&
                !x.Name.Equals("SubDep") &&
                !x.Name.Equals("LT")).ToList();

            List<SelectListItem> _menuList = new List<SelectListItem>();
            _menuList.Add(new SelectListItem
            {
                Text = "- Create New -",
                Value = "0"
            });

            foreach (var data in references)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Purpose,
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownHolidayType(IReferenceAppService service)
        {
            string dataList = service.GetDetailAll(ReferenceEnum.HolidayType, true);
            List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Description,
                    Value = data.Code,
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownFacilityDepartment(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string reference = service.GetBy("Name", "FacilityDepartment", true);
            ReferenceModel refModel = reference.DeserializeToReference();
            if (refModel != null)
            {
                string dataList = service.FindDetailBy("ReferenceID", refModel.ID, true);
                List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();

                foreach (var data in dataModelList)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = data.Description,
                        Value = data.ID.ToString()
                    });
                }
            }


            return _menuList;
        }

        public static List<SelectListItem> BindDropDownBrand(IReferenceAppService service)
        {
            string dataList = service.GetDetailAll(ReferenceEnum.Brand, true);
            List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Code,
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownBrandBlend(IBrandAppService brandService, IBlendAppService blendService)
        {
            string dataList = brandService.GetAll();
            List<BrandModel> dataModelList = dataList.DeserializeToBrandList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Code,
                    Value = data.ID.ToString()
                });
            }

            dataList = blendService.GetAll();
            List<BlendModel> blendModelList = dataList.DeserializeToBlendList();

            foreach (var data in blendModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Code,
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownGroupTypeCode(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string dataList = service.GetDetailAll(ReferenceEnum.Group, true);
            List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Code,
                    Value = data.Code
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownMachineBrandDesc(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string reference = service.GetBy("Name", "MB", true);
            ReferenceModel refModel = reference.DeserializeToReference();
            if (refModel != null)
            {
                string dataList = service.FindDetailBy("ReferenceID", refModel.ID, true);
                List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();

                foreach (var data in dataModelList)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = data.Description,
                        Value = data.Description
                    });
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownLegalEntityDesc(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string reference = service.GetBy("Name", "LE", true);
            ReferenceModel refModel = reference.DeserializeToReference();
            if (refModel != null)
            {
                string dataList = service.FindDetailBy("ReferenceID", refModel.ID, true);
                List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();

                foreach (var data in dataModelList)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = data.Description,
                        Value = data.Description
                    });
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownArea(IReferenceAppService service)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            string reference = service.GetBy("Name", "Area", true);
            ReferenceModel refModel = reference.DeserializeToReference();
            if (refModel != null)
            {
                string dataList = service.FindDetailBy("ReferenceID", refModel.ID, true);
                List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();

                foreach (var data in dataModelList)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = data.Description,
                        Value = data.Description
                    });
                }
            }

            return _menuList;
        }

        public static List<SelectListItem> BindDropDownWppType(IReferenceAppService service)
        {
            string dataList = service.GetDetailAll(ReferenceEnum.WppType, true);
            List<ReferenceDetailModel> dataModelList = dataList.DeserializeToRefDetailList();
            List<SelectListItem> _menuList = new List<SelectListItem>();

            foreach (var data in dataModelList)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = data.Description,
                    Value = data.ID.ToString()
                });
            }

            return _menuList;
        }

        public static List<SelectListItem> GetProductionCenterInIndonesia(ILocationAppService _locationAppService, IReferenceAppService _referenceAppService, bool isIncludePleaseSelect = true, long pcID = 0)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.ProdCenter).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refPC.DeserializeToRefDetailList();

            if (isIncludePleaseSelect)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "- Please Select -",
                    Value = "0"
                });
            }

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    if (pcID == 0)
                    {
                        _menuList.Add(new SelectListItem
                        {
                            Text = text.Code + " - " + text.Description,
                            Value = item.ID.ToString()
                        });
                    }
                    else
                    {
                        if (item.ID == pcID)
                        {
                            _menuList.Add(new SelectListItem
                            {
                                Text = text.Code + " - " + text.Description,
                                Value = item.ID.ToString()
                            });
                        }
                    }
                }
            }

            return _menuList;
        }

        public static MultiSelectList GetMultiProductionCenterInIndonesia(ILocationAppService _locationAppService, IReferenceAppService _referenceAppService, bool isIncludePleaseSelect = true)
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();
            string pcs = _locationAppService.FindBy("ParentCode", Constants.DEFAULT_COUNTRY, true);
            List<LocationModel> locModelList = pcs.DeserializeToLocationList().OrderBy(x => x.Code).ToList();

            string refPC = _referenceAppService.FindDetailBy("ReferenceID", ((int)ReferenceEnum.ProdCenter).ToString(), true);
            List<ReferenceDetailModel> refLocModelList = refPC.DeserializeToRefDetailList();

            if (isIncludePleaseSelect)
            {
                _menuList.Add(new SelectListItem
                {
                    Text = "- Please Select -",
                    Value = "0"
                });
            }

            foreach (var item in locModelList)
            {
                ReferenceDetailModel text = refLocModelList.Where(x => x.Code == item.Code).FirstOrDefault();
                if (text != null)
                {
                    _menuList.Add(new SelectListItem
                    {
                        Text = text.Code + " - " + text.Description,
                        Value = item.ID.ToString()
                    });
                }
            }

            MultiSelectList list = new MultiSelectList(_menuList, "Value", "Text");

            return list;
        }

        public static List<SelectListItem> BindDropDownGranularity()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "Daily",
                Value = "Daily"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Weekly",
                Value = "Weekly"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Shiftly",
                Value = "Shiftly"
            });

            return _menuList;
        }

        public static MultiSelectList BindDropDownMultiProduct()
        {
            List<SelectListItem> _menuList = new List<SelectListItem>();

            _menuList.Add(new SelectListItem
            {
                Text = "All",
                Value = "All"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Finish Goods",
                Value = "Finish"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Intermediate Product",
                Value = "Intermediate"
            });

            _menuList.Add(new SelectListItem
            {
                Text = "Primary",
                Value = "Primary"
            });

            MultiSelectList list = new MultiSelectList(_menuList, "Value", "Text");

            return list;
        }
    }
}