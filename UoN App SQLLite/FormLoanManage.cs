using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;

namespace UoN_App_SQLLite
{
    public partial class FormLoanManage : Form
    {
        public FormLoanManage()
        {
            InitializeComponent();
            HideAll();
        }

        private void FormLoanManage_Load(object sender, EventArgs e)
        {

        }

        public DataTable DBGetDataSet(string SQL)
        {
            // Requires using System.Data;

            string Connection = @"Data Source=\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";

            DataTable ReturnData = new DataTable();

            using (var c = new SQLiteConnection(Connection))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(SQL, c))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        try { ReturnData.Load(rdr); } catch { }
                    }
                }
            }

            return ReturnData;
        }

        public int FindColID(DataGridView DGV, string ColName)
        {
            int ColID = 0;

            //MessageBox.Show(DGV.Name.ToString());

            for (int i = 0; i <= DGV.ColumnCount - 1; i++)
            {
                // Get Column numbers
                if (DGV.Columns[i].HeaderText.ToString() == ColName) { ColID = i; }
            }
            return ColID;

        }

        void HideAll()
        {
            panelDeviceActivity.Visible = false;
        }

        private void buttonDeviceActivity_Click(object sender, EventArgs e)
        {
            HideAll();
            panelDeviceActivity.Visible = true;
            DGVDeviceList.DataSource = DBGetDataSet("SELECT * from LoanStock");


        }

        private void DGVDeviceList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string DeviceID = DGVDeviceList .CurrentRow.Cells[FindColID(DGVDeviceList,"Id")].Value.ToString();
            DGVDeviceActivity.DataSource = DBGetDataSet("Select * FROM LoanData WHERE Device = " + DeviceID + " ORDER BY CollectionDate Desc");
        }
    }
}
