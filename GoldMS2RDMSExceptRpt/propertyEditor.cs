using System;
using System.Windows.Forms;

namespace GoldMS2RDMSExceptRpt
{
    public partial class propertyEditor : Form
    {
        public propertyEditor()
        {
            InitializeComponent();
        }

        private void propertyEditor_Load(object sender, EventArgs e)
        {
            propertyGrid1.SelectedObject = new PropertyGridItems();
        }

        private void propertyEditor_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reset();
            propertyGrid1.Refresh();
        }
    }
}
