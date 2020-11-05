using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Fast.Web.Models
{
	public class MenuModel : BaseModel
	{
		public int DisplayOrder { get; set; }
		[Required]
		public string Name { get; set; }
		public string Description { get; set; }
		public string PageIcon { get; set; } = string.Empty;
		public string PageSlug { get; set; } = string.Empty;
        public string PageController { get; set; } = string.Empty;
        public string PageAction { get; set; } = string.Empty;
        public Nullable<long> ParentID { get; set; }
		public string ParentName { get; set; }
		public bool IsParent { get; set; }
		public bool IsTopMenu { get; set; }
	}

    public class MenuTreeModel: MenuModel
    {
        public bool IsParentTree { get; set; }
        public bool IsTopMenuTree { get; set; }
        public List<ParentMenuModel> Parents { get; set; }
        public MenuTreeModel()
        {
            Parents = new List<ParentMenuModel>();
        }
    }

    public class ParentMenuModel: MenuModel
    {
        public List<ChildMenuModel> Children { get; set; }
        public ParentMenuModel()
        {
            Children = new List<ChildMenuModel>();
        }
    }

    public class ChildMenuModel : MenuModel
    {
    }
}