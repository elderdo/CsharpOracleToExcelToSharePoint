using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace GoldMS2RDMSExceptRpt
{
    class PropertyGridItems
    {

        [Editor(@"System.Windows.Forms.Design.StringCollectionEditor," +
            "System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
            typeof(System.Drawing.Design.UITypeEditor))]
        [Category("Mail")]
        [DisplayName("Send To")]
        [Description("A collection of the recipients for the report")]
        public System.Collections.Specialized.StringCollection SendTo
        {
            get { return Properties.Settings.Default.sendTo; }
            set { Properties.Settings.Default.sendTo = value; }
        }

        [Editor(@"System.Windows.Forms.Design.StringCollectionEditor," +
            "System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
            typeof(System.Drawing.Design.UITypeEditor))]
        [Category("Mail")]
        [DisplayName("CC")]
        [Description("A collection of the CC'ed recipients for the report")]
        public System.Collections.Specialized.StringCollection CC
        {
            get { return Properties.Settings.Default.sendCC; }
            set { Properties.Settings.Default.sendCC = value; }
        }

        [Category("Forms")]
        [DisplayName("Show Display Grid")]
        [Description("Show the grid containing all the data retrieved for the report.")]
        public bool ShowDisplayGrid
        {
            get { return Properties.Settings.Default.displayGrid; }
            set { Properties.Settings.Default.displayGrid = value; }
        }

        [Category("Forms")]
        [DisplayName("Show Errors")]
        [Description("Show a message box containing any application errors.")]
        public bool ShowErrors
        {
            get { return Properties.Settings.Default.showMsgBox; }
            set { Properties.Settings.Default.showMsgBox = value; }
        }

        [Category("Forms")]
        [DisplayName("Confirm Duplicate")]
        [Description("Show a message box asking if you want to create another"
            + " report file, when one has already been created for the"
            + " report day and time.")]
        public bool ConfirmDuplicate
        {
            get { return Properties.Settings.Default.confirmDuplicate; }
            set { Properties.Settings.Default.confirmDuplicate = value; }
        }


        [Category("File")]
        [DisplayName("SharePoint UNC")]
        [Description("The location where the report resides")]
        public string SharePointUNC
        {
            get { return Properties.Settings.Default.sharePointUNC; }
            set { Properties.Settings.Default.sharePointUNC = value; }
        }

        [Category("File")]
        [DisplayName("Mail Message File")]
        [Description("The HTML file contain the body of the email"
        + " message sent by the application")]
        public string EmailMsg
        {
            get { return Properties.Settings.Default.emailMsg; }
            set { Properties.Settings.Default.emailMsg = value; }
        }

        [Category("File")]
        [DisplayName("The Report File Name")]
        [Description("The Excel file containing the report data.")]
        public string ExceptionReport
        {
            get { return Properties.Settings.Default.exceptRpt; }
            set { Properties.Settings.Default.exceptRpt = value; }
        }


        [Category("Execution")]
        [DisplayName("1. Run Day")]
        [Description("The day the report can be run")]
        public DayOfWeek RunDay
        {
            get { return Properties.Settings.Default.runDay; }
            set { Properties.Settings.Default.runDay = value; }
        }

        [Category("Execution")]
        [DisplayName("2. Start Time")]
        [Description("The earliest Mountain Time Hour the report can be started")]
        public int StartTime
        {
            get { return Properties.Settings.Default.startTime; }
            set { Properties.Settings.Default.startTime = value; }
        }

        [Category("Execution")]
        [DisplayName("3. End Time")]
        [Description("The latest Mountain Time Hour the report can be started")]
        public int EndTime
        {
            get { return Properties.Settings.Default.endTime; }

            set
            {
                if (StartTime < value)
                    Properties.Settings.Default.startTime = value;
                else
                    MessageBox.Show("End Time must be > Start Time", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
