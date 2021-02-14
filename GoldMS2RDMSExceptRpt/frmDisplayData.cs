using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Oracle.DataAccess.Client;
using System.Configuration;

namespace GoldMS2RDMSExceptRpt
{
    public partial class frmDisplayData : Form
    {
        OracleConnection conn;
        public frmDisplayData(OracleConnection conn)
        {
            this.conn = conn;
            InitializeComponent();
        }

        private void frmDisplayData_Load(object sender, EventArgs e)
        {
            ExceptRptData rptData = new ExceptRptData(conn);

            ExceptRpt rpt = new ExceptRpt();
            String fileName = "system_compare.xls";
            if (Properties.Settings.Default.exceptRpt != null)
                fileName = Properties.Settings.Default.exceptRpt;
            string sharePointUNC = @"\\vtr2b.web.boeing.com\ids\escmtools\Team Documents";
            if (Properties.Settings.Default.sharePointUNC != null)
                sharePointUNC = Properties.Settings.Default.sharePointUNC;

            rpt.createExceptRpt(rptData.DS,fileName);
            String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            SharePoint sp = new SharePoint();
            try
            {
                sp.upload(path + @"\" + fileName, sharePointUNC);
            }
            catch (BadExceptRptFile badExceptRptFile)
            {
                Logger.Warning(badExceptRptFile.Message, "frmDisplayData");
                if (Properties.Settings.Default.confirmDuplicate)
                {
                    DialogResult dialogResult = MessageBox.Show(
                        string.Format("{0}: Do you want to continue?",
                        badExceptRptFile.Message), "Create Error", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }

            }
            catch (Exception ex)
            {
                Logger.Error(ex.Message, "frmDisplayData");
                return;
            }

            
            Mail mail = new Mail();
            mail.SendMail();

            bool displayGrid = Properties.Settings.Default.displayGrid;
            if (displayGrid)
            {
                dataGridView1.DataSource = rptData.DS.Tables["ExceptRpt"].DefaultView;
                Int32 rowcnt = dataGridView1.RowCount - 1;
                lblRows.Text = (dataGridView1.CurrentRow.Index + 1).ToString() + " of " + rowcnt.ToString();
            }
            else
            {
                Application.Exit();
            }
        }
        private void dataGridView1_SelectionChanged(object sender, System.EventArgs e)
        {
            Int32 rowcnt = dataGridView1.RowCount - 1;
            if (dataGridView1.CurrentRow != null)
            {
                Int32 selectedRowNumber = dataGridView1.CurrentRow.Index + 1;
                lblRows.Text = selectedRowNumber.ToString() + " of " + rowcnt.ToString();
            }

        }

         protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData)
         {
             try
             {
                if (msg.WParam.ToInt32() == (int)Keys.Escape)
                {
                    this.Close();
                }
                else
                {
                    return base.ProcessCmdKey(ref msg, keyData);
                }
             }
             catch (Exception Ex )
             {
                 Logger.Error(Ex, "frmDisplayData");
             }
             return base.ProcessCmdKey(ref msg,keyData);
     }

         private void frmDisplayData_FormClosed(object sender, FormClosedEventArgs e)
         {
             Application.Exit();
         }
        /*
         private void InitializeComponent()
         {
            this.SuspendLayout();
            // 
            // frmDisplayData
            // 
            this.ClientSize = new System.Drawing.Size(311, 262);
            this.Name = "frmDisplayData";
            this.ResumeLayout(false);

         }
         */
    }
}