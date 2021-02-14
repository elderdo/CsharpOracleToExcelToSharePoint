using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Diagnostics;
using System.Configuration;


namespace GoldMS2RDMSExceptRpt
{
    public partial class Login : Form
    {
        OracleConnection myOracleConnection;
        string connectionString;
        private bool myIsAuthorized = false; 

        public Login()
        {
            InitializeComponent();
        }

        public bool IsAuthorized
        {
            get
            {
                return myIsAuthorized;
            }

        }

        public OracleConnection Connection
        {
            get
            {
                return myOracleConnection;
            }
        }
    
        private bool connectToDB(string connectionString)
        {
            myOracleConnection = new OracleConnection(connectionString);
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                myOracleConnection.Open();
                return true;
            }
            catch (OracleException ex)
            {
                if (ConfigurationManager.AppSettings["showMsgBox"].Equals("true"))
                    MessageBox.Show(ex.Message);
                else
                    Logger.Error(ex, "Login");
                return false;
            }

        }
        private bool isReportReady()
        {
            const string sql = "SELECT COUNT (*) FROM msms_process_log WHERE     module_name = 'MSMS.MSMS_SYS_COMPARE' AND TRIM (TO_CHAR (create_date, 'DAY')) = 'MONDAY' AND TRUNC (SYSDATE) BETWEEN TRUNC (create_date) AND TRUNC (create_date) + 6 AND log_text = 'Processing Ends'";
            try
            {
                using (OracleCommand ocmd = new OracleCommand(sql, myOracleConnection))
                {
                    return Convert.ToInt32(ocmd.ExecuteScalar()) > 0;
                }
            }
            catch (Exception e)
            {
                Logger.Error(e, "Login");
                return false;
            }
        }

        private void processData()
        {
            if (isReportReady())
            {
                frmDisplayData frm = new frmDisplayData(myOracleConnection);
                frm.ShowDialog();
            }
            else
            {
                bool showMsgBox = Properties.Settings.Default.showMsgBox;
                if (showMsgBox)
                    MessageBox.Show("MSMS.MSMS_SYSTEM_COMPARE has not completed");
                else
                    Logger.Error("MSMS.MSMS_SYSTEM_COMPARE has not completed","Login");
                Application.Exit();
            }

        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (textUserId.Text == ""
                || textPassword.Text == ""
                || comboHost.Text == "")
            {
                MessageBox.Show("You must enter an Oracle user id, password, and select an Oracle host");
            }
            else {
                string connectionString = "data source=" + comboHost.Text + ";user id="
                    + textUserId.Text + ";password=" + textPassword.Text;

                Cursor.Current = Cursors.WaitCursor;
                myIsAuthorized = connectToDB(connectionString);
                Cursor.Current = Cursors.Default;

                if (myIsAuthorized)
                {
                    this.Hide();
                    processData();
                }
            }
        }

        private bool appSettingsChanged;

 
        private void Login_Load(object sender, EventArgs e)
        {
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["Dev"].ConnectionString;
            if (connectionString != null)
            {
                myIsAuthorized = connectToDB(connectionString);

                if (myIsAuthorized)
                {
                    this.Hide();
                    processData();
                    return;
                }
                else
                {
                    Application.Exit();
                }

            }
            textUserId.Text = Environment.UserName;
            try
            {
                if (Application.UserAppDataRegistry.GetValue("HostString") != null)
                {
                    comboHost.Text =
                      Application.UserAppDataRegistry.GetValue(
                      "HostString").ToString();
                }
                if (Application.UserAppDataRegistry.GetValue("UserId") != null
                    && Application.UserAppDataRegistry.GetValue("UserId").ToString() != Environment.UserName)
                {
                    textUserId.Text =
                      Application.UserAppDataRegistry.GetValue(
                      "UserId").ToString();
                }
                
            }
            catch (Exception ex)
            {
                Logger.Error(ex,"Login");
            }           
        }

        private void comboHost_SelectedIndexChanged(object sender, EventArgs e)
        {
            appSettingsChanged = true;
        }

        private void textUserId_TextChanged(object sender, EventArgs e)
        {
            appSettingsChanged = true;
        }

        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (appSettingsChanged)
            {
                try
                {
                    Application.UserAppDataRegistry.SetValue("HostString",
                      comboHost.Text);
                    if (textUserId.Text != Environment.UserName)
                    {
                        Application.UserAppDataRegistry.SetValue("UserId",
                          textUserId.Text);

                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex,"Login");
                }
            }

        }

        private void Login_Activated(object sender, EventArgs e)
        {
            textPassword.Focus();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}