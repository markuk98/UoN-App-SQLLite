using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UoN_App_SQLLite
{
    public partial class FormChannelLog : Form
    {
        public FormChannelLog()
        {
            InitializeComponent();

        }

        private void ChannelLog_Load(object sender, EventArgs e)
        {
            //this.MaximumSize = new System.Drawing.Size(650, 350);

            webUpdateLog.AllowWebBrowserDrop = false;
            webUpdateLog.Url = new System.Uri("file://"+ Application.StartupPath + @"\ChannelChangeLog.htm");

        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
