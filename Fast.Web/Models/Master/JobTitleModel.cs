using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace Fast.Web.Models
{
    public class JobTitleModel : BaseModel
    {
        public string RoleName { get; set; }
        public string Code { get; set; } = string.Empty;
        [Required]
        public string Title { get; set; }
    }

    public class JobTitleTreeModel : JobTitleModel
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public List<SelectListItem> AvailableTitles { get; set; }
        public List<SelectListItem> SelectedTitles { get; set; }
        public List<ParentJobTitleModel> Parents { get; set; }
        public JobTitleTreeModel()
        {
            Parents = new List<ParentJobTitleModel>();
        }
    }

    public class ParentJobTitleModel : RoleModel
    {
        public List<JobTitleModel> Children { get; set; }
        public ParentJobTitleModel()
        {
            Children = new List<JobTitleModel>();
        }
    }
}