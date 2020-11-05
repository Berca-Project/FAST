using Fast.WinClient.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace Fast.WinClient
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            User userAccount;

            try
            {
                //Validations 				
                if (string.IsNullOrEmpty(txtUserName.Text) || string.IsNullOrEmpty(txtPassword.Text))
                {
                    MessageBox.Show("Username dan Password harus diisi");
                    return;
                }

                //Post Request for Login
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(@"http://localhost:1820/api/token");
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "POST";

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    userAccount = new User { UserName = txtUserName.Text, Password = txtPassword.Text };
                    string NewUserString = JsonConvert.SerializeObject(userAccount, Formatting.None);
                    streamWriter.Write(NewUserString);
                    streamWriter.Flush();
                    streamWriter.Close();
                }

                //Gettting response
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                //Reading response
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream(), Encoding.ASCII))
                {
                    string result = streamReader.ReadToEnd();
                    if (result.Contains("token"))
                    {
                        dynamic response = (JObject)JsonConvert.DeserializeObject(result);
                        userAccount.Token = response["token"];
                        GlobalProp.CurrentUser = userAccount;

                        MainForm mainForm = new MainForm();
                        mainForm.FormClosed += new FormClosedEventHandler(mainForm_FormClosed);
                        //showing the main form
                        mainForm.Show();
                        this.Hide();
                    }
                    else
                    {
                        MessageBox.Show(result);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Username atau Password tidak valid");
            }
        }

        void mainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            // here you can do anything

            // we will close the application
            Application.Exit();
        }
    }
}
