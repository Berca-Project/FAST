using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace Fast.Web.Models
{
	public class CRUDModel
	{
		/// <summary>  
		/// Get all records from the DB  
		/// </summary>  
		/// <returns>Datatable</returns>  
		public DataTable GetAllRoles()
		{
			DataTable dt = new DataTable();		
			string strConString = ConfigurationManager.ConnectionStrings["AdoAppConn"].ConnectionString;

			using (SqlConnection con = new SqlConnection(strConString))
			{
				con.Open();
				SqlCommand cmd = new SqlCommand("Select * from roles", con);
				SqlDataAdapter da = new SqlDataAdapter(cmd);
				da.Fill(dt);
			}

			return dt;
		}
	}
}