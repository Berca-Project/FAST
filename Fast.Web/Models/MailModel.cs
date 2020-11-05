namespace Fast.Web.Models
{
	public class MailModel
	{
		public string From { get; set; }
		public string To { get; set; }
		public string Subject { get; set; }
		public string Body { get; set; }
		public string SMTP { get; set; }
		public string Username { get; set; }
		public string Password { get; set; }

	}
}