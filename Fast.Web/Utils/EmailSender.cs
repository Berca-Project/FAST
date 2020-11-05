using System;
using System.Configuration;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using Fast.Application.Interfaces;

namespace Fast.Web.Utils
{
	public static class EmailSender
	{

		#region ::Email::
		public static async Task<bool> SendEmailAsync(string email, string msg, string dir, string subject = "")
		{
			// Initialization.
			bool isSend = false;

			try
			{
				// Initialization.
				var body = msg;
				var message = new MailMessage();

				// Settings.
				message.To.Add(new MailAddress(email));
				message.From = new MailAddress(ConfigurationManager.AppSettings["SmtpEmailFromAddress"]);
				message.Subject = !string.IsNullOrEmpty(subject) ? subject : ConfigurationManager.AppSettings["SmtpEmailDefaultSubject"];
				message.Body = body;
				message.IsBodyHtml = true;

				using (var smtp = new SmtpClient())
				{
					// Settings.
					var credential = new NetworkCredential
					{
						UserName = ConfigurationManager.AppSettings["SmtpEmailFromAddress"],
						Password = ConfigurationManager.AppSettings["SmtpEmailFromPassword"]
					};

					// Settings.					
					smtp.Host = ConfigurationManager.AppSettings["SmtpEmailHost"];
					smtp.Port = Convert.ToInt32(ConfigurationManager.AppSettings["SmtpEmailPort"]);
					smtp.EnableSsl = false;
					smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
					smtp.UseDefaultCredentials = false;
					//smtp.Credentials = CredentialCache.DefaultNetworkCredentials;
					smtp.Credentials = credential;

					// Sending
					await smtp.SendMailAsync(message);

					// Settings.
					isSend = true;
				}
			}
			catch (Exception ex)
			{
				if (dir == "")
				{
					dir = AppDomain.CurrentDomain.BaseDirectory;
				}

				Helper.LogErrorMessage("Email Error Log: " + ex.GetAllMessages(), dir);
			}

			// info.
			return isSend;
		}
		#endregion

		#region ::Email LPH::
		public static async Task<bool> SendEmailLPH(string toemail, string tujuan, string judul, string dir = "")
		{
			try
			{
				// Initialization.
				string emailMsg = "";
				emailMsg = emailMsg + "<div style='width:500px; text-align:center; padding-bottom:10px;'><br />";
				emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";
				emailMsg = emailMsg + "<h2>LPH Approval Request</h2>";
				emailMsg = emailMsg + "<br /><br /><div>";
				emailMsg = emailMsg + "To: " + tujuan + "<br /><br />";
				emailMsg = emailMsg + "<p>There is LPH submission which need your concern to be approved/rejected</p>";
				emailMsg = emailMsg + "<a href='" + WebConstants.BASEURLKU + "/Approval'>View LPH Submission</a>";
				emailMsg = emailMsg + "<br /><br />";
				emailMsg = emailMsg + "Best Regards,";
				emailMsg = emailMsg + "</div><br /><br />";
				emailMsg = emailMsg + "FAST Notification";

				string emailSubject = ConfigurationManager.AppSettings["SmtpEmailDefaultSubject"];

				// Sending Email.
				await SendEmailAsync(toemail, emailMsg, dir, emailSubject);

				return true;
			}
			catch (Exception ex)
			{
				// Info
				Console.Write(ex);
			}

			// Info
			return false;
		}
		#endregion


		#region ::Email LPH PP::
		public static async Task<bool> SendEmailLPHPP(string toemail, string tujuan, string judul, string dir = "")
		{
			try
			{
				// Initialization.
				string emailMsg = "";
				emailMsg = emailMsg + "<div style='width:500px; text-align:center; padding-bottom:10px;'><br />";
				emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";
				emailMsg = emailMsg + "<h2>LPH Approval Request</h2>";
				emailMsg = emailMsg + "<br /><br /><div>";
				emailMsg = emailMsg + "To: " + tujuan + "<br /><br />";
				emailMsg = emailMsg + "<p>There is LPH submission which need your concern to be approved/rejected</p>";
				emailMsg = emailMsg + "<a href='" + WebConstants.BASEURLKU + "/ApprovalPrimary'>View LPH Submission</a>";
				emailMsg = emailMsg + "<br /><br />";
				emailMsg = emailMsg + "Best Regards,";
				emailMsg = emailMsg + "</div><br /><br />";
				emailMsg = emailMsg + "FAST Notification";

                string emailSubject = ConfigurationManager.AppSettings["SmtpEmailDefaultSubject"];

                // Sending Email.
                await SendEmailAsync(toemail, emailMsg, dir, emailSubject);

				return true;
			}
			catch (Exception ex)
			{
				// Info
				Console.Write(ex);
			}

			// Info
			return false;
		}
        #endregion

        #region ::Email LPH PP Submitted::
        public static async Task<bool> SendEmailLPHPPSubmit(string toemail, string Requestor, string TYPELPHnya, string LPHController, string LPHIDNya, string dir = "")
        {
            try
            {
                // Initialization.
                string emailMsg = "";
                // Approval - Notification to Supervisor
				emailMsg = emailMsg + "<div style='width:500px; padding-bottom:10px;'><br /><br />";
                emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";

                emailMsg = emailMsg + "<h3>This document from "+Requestor+" needs your approval. </h3>";
				emailMsg = emailMsg + "Click the LPH "+TYPELPHnya+" no <a href='" + WebConstants.BASEURLKU + "/" + LPHController + "/Approval?lphid=" + LPHIDNya + "'>" + LPHIDNya+"</a> to view the details.";
				emailMsg = emailMsg + " Please submit your response and give any comment if needed. Thank you";
				emailMsg = emailMsg + "<br /><br /><hr>";
                
				emailMsg = emailMsg + "<i><h3>Dokumen dari " + Requestor+ " memerlukan persetujuan anda. </h3>";
				emailMsg = emailMsg + "Klik LPH "+TYPELPHnya+" no <a href='" + WebConstants.BASEURLKU + "/" + LPHController + "/Approval?lphid=" + LPHIDNya + "'>" + LPHIDNya+"</a> untuk melihat detail.";
				emailMsg = emailMsg + " Mohon memberikan tanggapan anda beserta komentar jika diperlukan. Terima kasih.</i>";
				emailMsg = emailMsg + "<br /><br />";

                emailMsg = emailMsg + "</div><br /><br /><hr>";
				emailMsg = emailMsg + "This is an automated message. Please do not respond to this email.";
                
                string emailSubject = "FAST: LPH "+TYPELPHnya+" "+LPHIDNya+" needs your approval" ;

                // Sending Email.
                await SendEmailAsync(toemail, emailMsg, dir, emailSubject);

                return true;
            }
            catch (Exception ex)
            {
                // Info
                Console.Write(ex);
            }

            // Info
            return false;
        }
        #endregion
        #region ::Email LPH PP Revise::
        public static async Task<bool> SendEmailLPHPPRevise(string toemail, string Approver, string TYPELPHnya, string LPHIDNya, string LPHController, string dir = "")
        {
            try
            {
                // Initialization.
                string emailMsg = "";
                // Revision - Notification to Prodtech
                emailMsg = emailMsg + "<div style='width:500px; padding-bottom:10px;'><br /><br />";
                emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";

                emailMsg = emailMsg + "<h3>" + Approver + " has verified your document and found some revisions to be done. </h3>";
                emailMsg = emailMsg + "<p>Please check the comment to make the revisions likewise, then re-submit. </p>";
                emailMsg = emailMsg + "Click LPH " + TYPELPHnya + " no <a href='" + WebConstants.BASEURLKU + "/" + LPHController + "/Edit?lphid=" + LPHIDNya + "'>" + LPHIDNya + "</a> to revise. Thank you.";
                emailMsg = emailMsg + "<br /><br /><hr>";

                emailMsg = emailMsg + "<i><p>" + Approver + " telah memeriksa dokumen anda dan menemukan beberapa revisi yang perlu dilakukan. </p>";
                emailMsg = emailMsg + "<h3>Mohon periksa komentar untuk revisi seperti yang diminta, dan kirimkan kembali. </h3>";
                emailMsg = emailMsg + "Klik LPH " + TYPELPHnya + " no <a href='" + WebConstants.BASEURLKU + "/" + LPHController + "/Edit?lphid=" + LPHIDNya + "'>" + LPHIDNya + "</a> untuk memperbaikinya. Terima kasih.</i>";
                emailMsg = emailMsg + "<br /><br />";

                emailMsg = emailMsg + "</div><br /><br /><hr>";
                emailMsg = emailMsg + "This is an automated message. Please do not respond to this email.";

                string emailSubject = "FAST: LPH " + TYPELPHnya + " " + LPHIDNya + " needs your revision";

                // Sending Email.
                await SendEmailAsync(toemail, emailMsg, dir, emailSubject);

                return true;
            }
            catch (Exception ex)
            {
                // Info
                Console.Write(ex);
            }

            // Info
            return false;
        }
        #endregion
        #region ::Email LPH PP Complete::
        public static async Task<bool> SendEmailLPHPPApprove(string toemail, string Requestor, string TYPELPHnya, string LPHIDNya, string LPHController, string dir = "")
        {
            try
            {
                // Initialization.
                string emailMsg = "";
                // Complete - Notification to Prodtech/Supervisor

                emailMsg = emailMsg + "<div style='width:500px; padding-bottom:10px;'><br /><br />";
                emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";

                emailMsg = emailMsg + "<h3>Approval request from " + Requestor + " for LPH " + TYPELPHnya + " no " + LPHIDNya + " has been completed. </h3>";
                emailMsg = emailMsg + "Click here " + TYPELPHnya + " no <a href='" + WebConstants.BASEURLKU + "/" + LPHController + "/Submitted?lphid=" + LPHIDNya + "'>" + LPHIDNya + "</a> to see the detail. Thank you.";
                emailMsg = emailMsg + "<br /><br /><hr>";

                emailMsg = emailMsg + "<i><h3>Permintaan persetujuan dari " + Requestor + " for LPH " + TYPELPHnya + " no " + LPHIDNya + " telah diselesaikan. </h3>";
                emailMsg = emailMsg + "Klik disini " + TYPELPHnya + " no <a href='" + WebConstants.BASEURLKU + "/" + LPHController + "/Submitted?lphid=" + LPHIDNya + "'>" + LPHIDNya + "</a> untuk melihat detail. Terima kasih.</i>";
                emailMsg = emailMsg + "<br /><br />";

                emailMsg = emailMsg + "</div><br /><br /><hr>";
                emailMsg = emailMsg + "This is an automated message. Please do not respond to this email.";

                string emailSubject = "FAST: LPH " + TYPELPHnya + " " + LPHIDNya + " has been completed";


               // Sending Email.
               await SendEmailAsync(toemail, emailMsg, dir, emailSubject);

                return true;
            }
            catch (Exception ex)
            {
                // Info
                Console.Write(ex);
            }

            // Info
            return false;
        }
        #endregion
        #region ::Email Checklist::
        public static async Task<bool> SendEmailChecklist(string toemail, string tujuan, string menuTitle, string header, string requestor, long submitID, string dir = "")
		{
			try
			{
				// Initialization.
				string emailMsg = "";
				emailMsg = emailMsg + "<div style='width:500px; text-align:center; padding-bottom:10px;'><br />";
				emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";
				emailMsg = emailMsg + "<h2>Checklist Approval Request: \""+ menuTitle +"\"</h2>";
				emailMsg = emailMsg + "<br /><br /><div>";
				emailMsg = emailMsg + "Dear " + tujuan + "<br /><br />";
				emailMsg = emailMsg + "<p>The following \""+ header +"\" submitted by \""+requestor+ "\" needs your review, please <a href='" + WebConstants.BASEURLKU + "/Checklist/Submitted/" + submitID + "'>click here</a> to go to the document.</p>";
				emailMsg = emailMsg + "<p>Thank you for your attention and cooperation.</p>";
				emailMsg = emailMsg + "<br /><br />";
				emailMsg = emailMsg + "Best Regards,";
				emailMsg = emailMsg + "</div><br /><br />";
				emailMsg = emailMsg + "FAST Notification";

				string emailSubject = "[FAST] Checklist Approval Request: " + menuTitle;

				// Sending Email.
				await SendEmailAsync(toemail, emailMsg, dir, emailSubject);

				return true;
			}
			catch (Exception ex)
			{
				// Info
				Console.Write(ex);
			}

			return false;
		}

        public static async Task<bool> SendEmailChecklistSubmitter(string toemail, string submitter, string menuTitle, string header, string approver, long submitID, string status,string dir = "")
        {
            try
            {
                // Initialization.
                string emailMsg = "";
                emailMsg = emailMsg + "<div style='width:500px; text-align:center; padding-bottom:10px;'><br />";
                emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";
                emailMsg = emailMsg + "<h2>Checklist Approval Status for \"" + menuTitle + "\" is <b>"+status+"</b></h2>";
                emailMsg = emailMsg + "<br /><br /><div>";
                emailMsg = emailMsg + "Dear " + submitter + "<br /><br />";
                emailMsg = emailMsg + "<p>The following \"" + header + "\" has been reviewed by "+approver+" with status: <b>"+status+"</b>. ";
                
                if (status == "Need to be revised")
                    emailMsg = emailMsg + "Please <a href='" + WebConstants.BASEURLKU + "/Checklist/EditValue/" + submitID + "_0'>click here</a> to revise the document.</p>";
                else
                    emailMsg = emailMsg + "You can <a href='" + WebConstants.BASEURLKU + "/Checklist/ViewOnly/" + submitID + "_0'>click here</a> to go to the document.</p>";
                
                emailMsg = emailMsg + "<p>Thank you for your attention and cooperation.</p>";
                emailMsg = emailMsg + "<br /><br />";
                emailMsg = emailMsg + "Best Regards,";
                emailMsg = emailMsg + "</div><br /><br />";
                emailMsg = emailMsg + "FAST Notification";

                string emailSubject = "[FAST] Checklist Approval Status for " + menuTitle;

                // Sending Email.
                await SendEmailAsync(toemail, emailMsg, dir, emailSubject);

                return true;
            }
            catch (Exception ex)
            {
                // Info
                Console.Write(ex);
            }

            return false;
        }

        #endregion
        #region ::Email Report Shiftly Daily::
        public static async Task<bool> SendEmailReportShiftlyDaily(DateTime date, string prodcenter, string machine, string emailTo, string dir = "")
        {
            try
            {
                // Initialization.
                string emailMsg = "";
                // Approval - Notification to Supervisor
                emailMsg = emailMsg + "<div style='width:500px; padding-bottom:10px;'><br /><br />";
                emailMsg = emailMsg + "<img src='" + WebConstants.BASEURLKU + "/Content/theme/images/logo.jpg'></div>";

                emailMsg = emailMsg + "<h3>This is Report Shiftly and Daily. </h3>";
                emailMsg = emailMsg + "<a href='" + WebConstants.BASEURLKU + "/ShiftDailyUrlExcel/GenerateShiftlyDailyWithParam?key=S2F0YWthbmxhaCAoTXVoYW1tYWQpLCAnRGlhbGFoIEFsbGFoLCBZYW5nIE1haGEgRXNhJy4gQWxsYWggdGVtcGF0IG1lbWludGEgc2VnYWxhIHNlc3VhdHUuIChBbGxhaCkgdGlkYWsgYmVyYW5hayBkYW4gdGlkYWsgcHVsYSBkaXBlcmFuYWtrYW4uIERhbiB0aWRhayBhZGEgc2VzdWF0dSB5YW5nIHNldGFyYSBkZW5nYW4gRGlhLg==&date=" + date.ToString("yyyy-MM-dd") + "&prodcenter=" + prodcenter + "&machine=" + machine + "'>Download</a> here.";
                emailMsg = emailMsg + "Thank you.";
                emailMsg = emailMsg + "<br /><br /><hr>";

                emailMsg = emailMsg + "<i><h3>Ini merupakan Report Shiftly and Daily. </h3>";
                emailMsg = emailMsg + "<a href='" + WebConstants.BASEURLKU + "/ShiftDailyUrlExcel/GenerateShiftlyDailyWithParam?key=S2F0YWthbmxhaCAoTXVoYW1tYWQpLCAnRGlhbGFoIEFsbGFoLCBZYW5nIE1haGEgRXNhJy4gQWxsYWggdGVtcGF0IG1lbWludGEgc2VnYWxhIHNlc3VhdHUuIChBbGxhaCkgdGlkYWsgYmVyYW5hayBkYW4gdGlkYWsgcHVsYSBkaXBlcmFuYWtrYW4uIERhbiB0aWRhayBhZGEgc2VzdWF0dSB5YW5nIHNldGFyYSBkZW5nYW4gRGlhLg==&date=" + date.ToString("yyyy-MM-dd") + "&prodcenter=" + prodcenter + "&machine=" + machine + "'>Download</a> disini.";
                emailMsg = emailMsg + "Terima kasih.</i>";
                emailMsg = emailMsg + "<br /><br />";

                emailMsg = emailMsg + "</div><br /><br /><hr>";
                emailMsg = emailMsg + "This is an automated message. Please do not respond to this email.";

                string emailSubject = "FAST: Report Shiftly and Daily "+date.ToString("dd-MMM-yy");

                // Sending Email.
                await SendEmailAsync(emailTo, emailMsg, dir, emailSubject);

                return true;
            }
            catch (Exception ex)
            {
                // Info
                Console.Write(ex);
            }

            // Info
            return false;
        }
        #endregion
    }
}