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
using System.Net.Mail;
using System.IO;
using System.Collections;
using System.Diagnostics;
using Microsoft.Win32;
using System.Data.SqlClient;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;






namespace UoN_App_SQLLite
{

    public partial class FormSEANS : Form
    {
        ROSDBQ ROSDBQuery = new ROSDBQ();
        EncryptClass ENC = new EncryptClass();

        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;


        Color currentColor = Color.Green;
        bool SEANSSave = true;
        string OperaterID = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

        int ChaseTicketIDCol = 0;
        int ChaseTicketRef = 0;
        int ChaseEmailIDCol = 0;
        int ChaseStatusIDCol = 0;
        int ChaseInfoIDCol = 0;
        int ChaseUserCol = 0;
        int ChaseTicketTitle = 0;
        int ChaseReminderCol = 0;



        public FormSEANS()
        {
            SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;  // Subscribe to the SessionSwitch event

            this.components = new System.ComponentModel.Container();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.menuItem1 = new System.Windows.Forms.MenuItem();

            // Initialize contextMenu1
            this.contextMenu1.MenuItems.AddRange(
                        new System.Windows.Forms.MenuItem[] { this.menuItem1 });

            // Initialize menuItem1
            this.menuItem1.Index = 0;
            this.menuItem1.Text = "E&xit";
            this.menuItem1.Click += new System.EventHandler(this.menuItem_Click);

            // Set up how the form should be displayed.
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Text = "Notify Icon Example";

            // Create the NotifyIcon.
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);

            // The Icon property sets the icon that will appear
            // in the systray for this application.
            notifyIcon.Icon = new Icon("Resources\\Phone_Icon.ico");

            // The ContextMenu property sets the menu that will
            // appear when the systray icon is right clicked.
            notifyIcon.ContextMenu = this.contextMenu1;

            // The Text property sets the text that will be displayed,
            // in a tooltip, when the mouse hovers over the systray icon.
            notifyIcon.Text = "Form1 (NotifyIcon example)";
            notifyIcon.Visible = false;

            // Handle the DoubleClick event to activate the form.
            notifyIcon.DoubleClick += new System.EventHandler(this.notifyIcon_DoubleClick);

            InitializeComponent();
            {

                Screen screen = Screen.PrimaryScreen;
                int S_width = screen.Bounds.Width;
                int S_height = screen.Bounds.Height;

                int NewWidth = (375 + (Properties.Settings.Default.FormWidth * 5));

                this.MaximumSize = new System.Drawing.Size(NewWidth, S_height);
                dataGridView1.Height = ClientRectangle.Height - 255;
                LoanstabControl.Height = ClientRectangle.Height - 75;

                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;


                ButtonVersion.Text = String.Format("Version {0}", version);

                //var version2 = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;

                //string version2 = System.Windows.Forms.Application.ProductVersion;
                //ButtonVersion.Text = String.Format("Version {0}", version2);

                //ButtonVersion.Text = String.Format("Version {0}", version2);

                //string conStringDatosUsuarios = Application.ExecutablePath + "UoNSeansData.db3";
                bool DBFile = System.IO.File.Exists(Application.StartupPath + @"\UoNSeansData.db3");

                if (DBFile)
                {
                    //ButtonSave.Visible = true;
                    //MessageBox.Show("Yes");

                }

                //ButtonSave.Visible = true;




                ReadData();
                ReadTodaysBookings();
                PhoneUpdate();

                Properties.Settings.Default.ChannelLog = true;
                String UserID = Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name);

                try
                {
                    //*/
                    string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                    using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                    {
                        using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                        {
                            Conn.Open();

                            //MessageBox.Show(Convert.ToString(System.Security.Principal.WindowsIdentity.GetCurrent().Name));
                            cmd.CommandText = "select * from admins";

                            using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {

                                    //MessageBox.Show(UserID + " vs " + Convert.ToString(reader.GetValue(reader.GetOrdinal("uniID"))));
                                    if (string.Equals(UserID, Convert.ToString(reader.GetValue(reader.GetOrdinal("uniID"))), StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        //MessageBox.Show("" + reader.GetValue(reader.GetOrdinal("Permissions")));

                                        Properties.Settings.Default.Permissions = Convert.ToString(reader.GetValue(reader.GetOrdinal("Permissions")));
                                        Properties.Settings.Default.Save();


                                        //MessageBox.Show("" + Properties.Settings.Default.Permissions);

                                    }
                                }


                            }
                        }
                    }


                    if (Convert.ToInt32(Properties.Settings.Default.Permissions) >= 0)
                    {

                    }
                }



                catch { }
                //*/



            }
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            bool A = true;
            if (A == true)
            {
                hScrollBar1.Value = Properties.Settings.Default.FormWidth;
                textBoxPhoneNumber.Text = Properties.Settings.Default.PhoneNumber;

                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                var versionstring = String.Format("Version {0}", version);
                Properties.Settings.Default.Save();
                //MessageBox.Show("" + Properties.Settings.Default.ChannelLog);

                if (Properties.Settings.Default.LastVersion != versionstring)
                {
                    Properties.Settings.Default.LastVersion = versionstring;
                    Properties.Settings.Default.Save();

                    FormChannelLog frm = new FormChannelLog();
                    frm.ShowDialog(this);

                    frm.Dispose();



                }
                //MessageBox.Show("" + Properties.Settings.Default.ChannelLog);
            }
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
        void ReadDB()
        {
            string conStringDatosUsuarios = @"\UoN App\UoNSeansData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    Conn.Open();
                    //ORDER BY column1, column2, ... ASC|DESC
                    cmd.CommandText = "SELECT * FROM UserData ORDER BY Id ASC";

                    using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            dataGridView1.Rows.Add(new object[]
                            {
                                    reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("s")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("e")),
                                    reader.GetValue(reader.GetOrdinal("a")),
                                    reader.GetValue(reader.GetOrdinal("n")),
                                    reader.GetValue(reader.GetOrdinal("Sol")),
                                    reader.GetValue(reader.GetOrdinal("Id"))
                            });
                        }
                    }

                    Conn.Close();



                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = (5 * 1000); // 5 secs
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();
        }

        private void ButtonNA2_Click(object sender, EventArgs e)
        {
            textBoxE.Text = "N/A";
        }

        private void ButtonNA3_Click(object sender, EventArgs e)
        {
            textBoxA.Text = "N/A";
        }

        private void ButtonNA4_Click(object sender, EventArgs e)
        {
            textBoxN.Text = "N/A";
        }

        private void ButtonClear_Click(object sender, EventArgs e)
        {
            textBoxS.Text = "";
            textBoxE.Text = "";
            textBoxA.Text = "";
            textBoxN.Text = "";
            textBoxSol.Text = "";
            textBoxS.Focus();
        }

        void CopyToClipBoard()
        {
            try
            {

                var TextString = ("S (Situation): " + textBoxS.Text + Environment.NewLine +
                          "E (Escalation): " + textBoxE.Text + Environment.NewLine +
                          "A (Action): " + textBoxA.Text + Environment.NewLine +
                          "N (Next Step): " + textBoxN.Text + Environment.NewLine +
                          "S (Solution): " + textBoxSol.Text + Environment.NewLine +
                          Environment.NewLine + "Notes: ");

                Clipboard.SetText(TextString);
            }
            catch { }
        }

        void SaveToHistory()
        {

            CopyToClipBoard();
            var NotNew = false;
            if (textBoxS.Text == "")
            {
                //MessageBox.Show("Empty");
            }
            else
            {


                try
                {
                    // Check if entry has been recalled from history
                    //string newstr = str.Replace("tag", "newtag");
                    if (textBoxS.Text == dataGridView1.CurrentRow.Cells[1].Value.ToString() &&
                        textBoxE.Text == dataGridView1.CurrentRow.Cells[2].Value.ToString() &&
                        textBoxA.Text == dataGridView1.CurrentRow.Cells[3].Value.ToString() &&
                        textBoxN.Text == dataGridView1.CurrentRow.Cells[4].Value.ToString() &&
                        textBoxSol.Text == dataGridView1.CurrentRow.Cells[5].Value.ToString()
                        )
                    {

                        //MessageBox.Show("True");
                        NotNew = true;
                    }
                }

                catch { }



                try
                {



                    if (NotNew == false)
                    {
                        NewEntry();
                    }

                }

                catch (Exception err)
                {
                    MessageBox.Show("Error..." + err);
                }
            }
        }

        void SaveToSave()
        {
            try
            {

                var TextString = ("S (Situation):" + textBoxS.Text + Environment.NewLine +
                          "E (Escalation):" + textBoxE.Text + Environment.NewLine +
                          "A (Action):" + textBoxA.Text + Environment.NewLine +
                          "N (Next Step):" + textBoxN.Text + Environment.NewLine +
                          "S (Solution):" + textBoxSol.Text + Environment.NewLine +
                          Environment.NewLine + "Notes: ");

                Clipboard.SetText(TextString);
            }
            catch { }

            var NotNew = false;
            if (textBoxS.Text == "")
            {

            }
            else
            {


                try
                {
                    // Check if entry has been recalled from history
                    if (textBoxS.Text == dataGridView1.CurrentRow.Cells[0].Value.ToString() &&
                        textBoxE.Text == dataGridView1.CurrentRow.Cells[1].Value.ToString() &&
                        textBoxA.Text == dataGridView1.CurrentRow.Cells[2].Value.ToString() &&
                        textBoxN.Text == dataGridView1.CurrentRow.Cells[3].Value.ToString() &&
                        textBoxSol.Text == dataGridView1.CurrentRow.Cells[4].Value.ToString()
                        )
                    {

                        //MessageBox.Show("True");
                        NotNew = true;
                    }
                }

                catch { }



                try
                {



                    if (NotNew == false)
                    {

                    }

                }

                catch (Exception err)
                {
                    MessageBox.Show("Error..." + err);
                }
            }
        }

        private void ButtonCopy_Click(object sender, EventArgs e)
        {
            SaveToHistory();

        }

        void NewEntry()
        {

            string conStringDatosUsuarios = @"\UoN App\UoNSeansData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    //MessageBox.Show("Hi");
                    Conn.Open();
                    textBoxS.Text = textBoxS.Text.Replace("'", "''");
                    textBoxE.Text = textBoxE.Text.Replace("'", "''");
                    textBoxA.Text = textBoxA.Text.Replace("'", "''");
                    textBoxN.Text = textBoxN.Text.Replace("'", "''");
                    textBoxSol.Text = textBoxSol.Text.Replace("'", "''");

                    if (SEANSSave == true)
                    {
                        //MessageBox.Show("INSERT INTO UserData(s,e,a,n,Sol) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')");
                        cmd.CommandText = "INSERT INTO UserData(s,e,a,n,Sol) values('" + textBoxS.Text + "','" + textBoxE.Text + "','" + textBoxA.Text + "','" + textBoxN.Text + "','" + textBoxSol.Text + "')";
                        cmd.ExecuteNonQuery();
                    }
                    dataGridView1.Rows.Clear();

                    Conn.Close();

                    ReadData();

                    textBoxS.Text = textBoxS.Text.Replace("''", "'");
                    textBoxE.Text = textBoxE.Text.Replace("''", "'");
                    textBoxA.Text = textBoxA.Text.Replace("''", "'");
                    textBoxN.Text = textBoxN.Text.Replace("''", "'");
                    textBoxSol.Text = textBoxSol.Text.Replace("''", "'");
                }
            }
        }

        void ReadData()
        {
            var DBSearchString = "";

            if (textBoxS.Text != "")
            {
                var SearchString = textBoxS.Text.Replace("'", "''");
                DBSearchString = " WHERE s LIKE '%" + SearchString + "%'";

            }
            dataGridView1.Rows.Clear();
            //Yangtzee/shared4/INS/IS_Info/A_Service%20Strategy/Teams/Service%20Delivery/Audio%20Visual%20and%20Events/Loan%20Database/LoanDB.mdf
            string conStringDatosUsuarios = @"\UoN App\UoNSeansData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    Conn.Open();
                    //ORDER BY column1, column2, ... ASC|DESC
                    cmd.CommandText = "SELECT * FROM UserData" + DBSearchString + " ORDER BY Id DESC";

                    dataGridView1.Rows.Clear();

                    using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            dataGridView1.Rows.Add(new object[]
                            {
                            reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("s")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("e")),
                                    reader.GetValue(reader.GetOrdinal("a")),
                                    reader.GetValue(reader.GetOrdinal("n")),
                                    reader.GetValue(reader.GetOrdinal("Sol")),
                                    reader.GetValue(reader.GetOrdinal("Id"))
                            });
                        }
                    }

                    Conn.Close();
                }
            }
        }

        void DeleteEntry()
        {

            try
            {
                if (dataGridView1.CurrentRow.Cells[0].Value.ToString() != "")
                {
                    //MessageBox.Show(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    DialogResult dialogResult = MessageBox.Show("Delete " + dataGridView1.CurrentRow.Cells[1].Value.ToString() + " from your history?" + Environment.NewLine + " It can NOT be undone (yet)!", "Confirm Delete", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                        string conStringDatosUsuarios = @"\UoN App\UoNSeansData.db3";
                        using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                        {
                            using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                            {

                                Conn.Open();


                                //MessageBox.Show("INSERT INTO UserData(s,e,a,n,Sol) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')");
                                // DELETE FROM table_name WHERE condition;
                                cmd.CommandText = "DELETE FROM UserData WHERE Id='" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "SELECT * FROM UserData";

                                Conn.Close();

                                ReadData();
                            }
                        }
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        //do something else
                    }
                }
            }
            catch { }



        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            LoadFromGridView();
        }

        void LoadFromGridView()
        {
            try { textBoxS.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString(); } catch { }
            try { textBoxE.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString(); } catch { }
            try { textBoxA.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString(); } catch { }
            try { textBoxN.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString(); } catch { }
            try { textBoxSol.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString(); } catch { }

            CopyToClipBoard();

        }

        private void deleteRecordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteEntry();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ReadData();
            }
            catch { }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            //dataGridView1.Height = Form1.Height - 230;
            dataGridView1.Height = ClientRectangle.Height - 230;
            LoanstabControl.Height = ClientRectangle.Height - 70;



        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            textBoxS.SelectAll();
            textBoxS.Focus();
        }

        private void textBox2_DoubleClick(object sender, EventArgs e)
        {
            textBoxE.SelectAll();
            textBoxE.Focus();
        }

        private void textBox3_DoubleClick(object sender, EventArgs e)
        {
            textBoxA.SelectAll();
            textBoxA.Focus();
        }

        private void textBox4_DoubleClick(object sender, EventArgs e)
        {
            textBoxA.SelectAll();
            textBoxA.Focus();
        }

        private void textBox5_DoubleClick(object sender, EventArgs e)
        {
            textBoxSol.SelectAll();
            textBoxSol.Focus();
        }

        private void ButtonVersion_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(Application.ExecutablePath);
            //MessageBox.Show(Properties.Settings.Default.LastVersion);
            //Clipboard.SetText(Application.StartupPath);
            FormChannelLog frm = new FormChannelLog();
            frm.ShowDialog(this);

            frm.Dispose();

        }

        private void ButtonSave_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyValue == 13) textBoxE.Focus();

            if (e.KeyCode == Keys.Enter && Control.ModifierKeys == Keys.Control)
            {
                e.Handled = false;
                LoadFromGridView();
            }

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13) textBoxA.Focus();
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13) textBoxN.Focus();
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13) textBoxSol.Focus();
        }

        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13) ButtonCopy.Focus();
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            //MessageBox.Show("" + e.KeyValue);
            if (e.KeyValue == 46)
            {
                ButtonCopy.Focus();
                DeleteEntry();
            }
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {


            if (e.Button == MouseButtons.Right)
            {
                try
                {


                    var hti = dataGridView1.HitTest(e.X, e.Y);
                    dataGridView1.ClearSelection();
                    dataGridView1.CurrentCell = dataGridView1.Rows[hti.RowIndex].Cells[1];

                    dataGridView1.Rows[hti.RowIndex].Selected = true;
                }
                catch { }
            }
        }

        private void timer_Tick(object sender, EventArgs e)
        {

            CheckPhones();

            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == false)
            {
                return;
            }
            String Hostname = System.Windows.Forms.SystemInformation.ComputerName;

            string conStringDatosUsuarios = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";

            String NeedsUpdate = (DBSingleRead("Select GetUpdates from Updates where HostName = '" + Hostname + "'", conStringDatosUsuarios));


            if (NeedsUpdate == "")
            {
                DBWright("INSERT INTO Updates (HostName, GetUpdates) values ('" + Hostname + "', 'false')", conStringDatosUsuarios);
            }
            else if (NeedsUpdate == "false")
            {
                return;
            }
            //MessageBox.Show("UpdateDefaultButton");
            ReadTodaysBookings();
            PhoneUpdate();

            DBWright("UPDATE Updates SET GetUpdates = 'false' WHERE HostName = '" + Hostname + "'", conStringDatosUsuarios);


        }

        void ReadTodaysBookings()
        {
            try
            {
                //MessageBox.Show(LoanLocationsView1.CurrentRow.Cells[1].Value.ToString());
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                        //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)
                        //MessageBox.Show(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (NOT (LoanData.ID IN (SELECT Id FROM LoanData WHERE returned = collected AND returned not null))) AND CollectionDate > '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 01:00:00' and CollectionDate < '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 23:00:00'"; //AND LoanData.Collected = 0
                        cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id, Customer.Firstname FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (Collectiondate > '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 08:00:00' And Collectiondate < '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 18:00:00' And Collected is null And Returned is null) or (Collectiondate < '" + DateTime.Now.ToString("yyyy-MM-dd") + " 08:00:00' And Collected is null And Returned is null) ORDER BY Collectiondate ASC";


                        CollectsdataGridView.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                CollectsdataGridView.Rows.Add(new object[]
                                {
                            reader.GetValue(7) + " " + reader.GetValue(0),  // U can use column index
                            reader.GetValue(1),        //reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                            reader.GetValue(2),        //reader.GetValue(reader.GetOrdinal("Description")),
                            reader.GetValue(3),        //reader.GetValue(reader.GetOrdinal("Location")),
                            reader.GetValue(4),
                            reader.GetValue(5),
                            reader.GetValue(6)//reader.GetValue(reader.GetOrdinal("CreatedBy")),
                                    //reader.GetValue(reader.GetOrdinal("CollectionDate"))


                                });
                            }
                        }

                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Returned, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (NOT (LoanData.ID IN (SELECT Id FROM LoanData WHERE returned = collected AND returned not null))) AND ReturnDate > '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 01:00:00' and ReturnDate < '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 23:00:00'"; //AND LoanData.Returned = 0
                        cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.ReturnDate, LoanData.Returned, LoanData.Id, Customer.Firstname FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE Customer.ID != 1 and (Returndate > '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 08:00:00' And Returndate < '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 18:00:00' And Returned is null) or (Returndate < '" + DateTime.Now.ToString("yyyy-MM-dd") + " 08:00:00' And Returned is null) ORDER BY Returndate ASC";


                        ReturnsdataGridView.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ReturnsdataGridView.Rows.Add(new object[]
                                {
                            reader.GetValue(7) + " " + reader.GetValue(0),  // U can use column index
                            reader.GetValue(1),        //reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                            reader.GetValue(2),        //reader.GetValue(reader.GetOrdinal("Description")),
                            reader.GetValue(3),        //reader.GetValue(reader.GetOrdinal("Location")),
                            reader.GetValue(4),
                            reader.GetValue(5),
                            reader.GetValue(6)        //reader.GetValue(reader.GetOrdinal("CreatedBy")),
                                    //reader.GetValue(reader.GetOrdinal("CollectionDate"))


                                });
                            }
                        }

                        String NowTime0 = Convert.ToString(DateTime.Now.AddDays(0).ToString("yyyy-MM-dd") + " 17:00");
                        String NowTime7 = Convert.ToString(DateTime.Now.AddDays(7).ToString("yyyy-MM-dd") + " 00:00");
                        //MessageBox.Show(NowTime);
                        cmd.CommandText = @"
select 

AssetID,

LCustomer.FirstName,
LCustomer.Surname,
LCustomer.Telephone,
LID,
LReturnDate,
NCollectionDate,
NID,
NCustomer.FirstName,
NCustomer.Surname,
NCustomer.Telephone


from LoanStock 

left join 
(Select loandata.ID as LID, loandata.Device as LDevice, Customer as LateCustomer, ReturnDate as LReturnDate From loandata 
where collected is not null 
and returned is null
and LoanData.returndate < '" + NowTime0 + @"')
on LDevice = LoanStock.ID

Left Join Customer as LCustomer on LCustomer.ID = LateCustomer 

left join 
(Select loandata.ID as NID, loandata.Device as NDevice, Customer as NextCustomer, CollectionDate as NCollectionDate From loandata 
where collected is null 
and returned is null
and LoanData.collectiondate < '" + NowTime7 + @"')
on NDevice = LoanStock.ID


Left Join Customer as NCustomer on NCustomer.ID = NextCustomer

Where LID is not null and NID is not null and LCustomer.ID != NCustomer.ID order by NCollectionDate asc
";


                        DangerLoandfataGridView.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {

                                DangerLoandfataGridView.Rows.Add(new object[]
                                {

                            reader.GetValue(0),  // U can use column index
                            reader.GetValue(1)+ ", " + reader.GetValue(2),        //reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                            reader.GetValue(3),        //reader.GetValue(reader.GetOrdinal("Description")),
                            reader.GetValue(4),        //reader.GetValue(reader.GetOrdinal("Location")),
                            reader.GetValue(5),
                            reader.GetValue(6),
                            reader.GetValue(7),
                            reader.GetValue(8) + ", " + reader.GetValue(9),//reader.GetValue(reader.GetOrdinal("CreatedBy")),
                            reader.GetValue(10)       //reader.GetValue(reader.GetOrdinal("CollectionDate"))


                                });
                            }
                        }



                        //Combo box for device return

                        cmd.CommandText = @"
                                            Select AssetID from LoanData 
                                            left join Loanstock on Loandata.Device = LoanStock.ID
                                            where Collected is not null and returned is null
                                            order by AssetID Asc
                                           ";


                        DevicesOutcomboBox.Items.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {

                                DevicesOutcomboBox.Items.Add(reader.GetValue(reader.GetOrdinal("AssetID")).ToString());

                            }
                        }
                    }
                    Conn.Close();

                    foreach (DataGridViewRow row in CollectsdataGridView.Rows)
                        if (Convert.ToDateTime(row.Cells[4].Value.ToString()) < Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 01:00"))
                        {
                            row.DefaultCellStyle.BackColor = Color.Orange;
                        }

                    foreach (DataGridViewRow row in ReturnsdataGridView.Rows)
                        if (Convert.ToDateTime(row.Cells[4].Value.ToString()) < Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 01:00"))
                        {
                            row.DefaultCellStyle.BackColor = Color.Orange;
                        }

                    if (DangerLoandfataGridView.RowCount == 0)
                    {
                        LoanstabControl.TabPages[2].Text = "";
                        TabFlashtimer.Stop();
                    }
                    else
                    {
                        //LoanstabControl.TabPages[2].Text = "ATTENTION!!";
                        TabFlashtimer.Start();
                    }


                }
            }
            catch
            {
                //MessageBox.Show("Error..." + err);
            }

            //CheckReminder();
        }

        private void PhoneUpdate()
        {
            try
            {


                buttonPhone1.BackColor = Color.Gainsboro;
                buttonPhone1.ForeColor = Color.Gainsboro;
                buttonPhone2.BackColor = Color.Gainsboro;
                buttonPhone2.ForeColor = Color.Gainsboro;
                buttonPhone3.BackColor = Color.Gainsboro;
                buttonPhone3.ForeColor = Color.Gainsboro;
                buttonPhone4.BackColor = Color.Gainsboro;
                buttonPhone4.ForeColor = Color.Gainsboro;
                buttonPhone5.BackColor = Color.Gainsboro;
                buttonPhone5.ForeColor = Color.Gainsboro;
                buttonPhone6.BackColor = Color.Gainsboro;
                buttonPhone6.ForeColor = Color.Gainsboro;


                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\PhoneDB.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                        //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)
                        //MessageBox.Show(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (NOT (LoanData.ID IN (SELECT Id FROM LoanData WHERE returned = collected AND returned not null))) AND CollectionDate > '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 01:00:00' and CollectionDate < '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 23:00:00'"; //AND LoanData.Collected = 0
                        cmd.CommandText = "SELECT UniID, PhoneNumber FROM Activeusers WHERE Logout is null";

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            button6.BackColor = Color.Goldenrod;
                            button6.Text = "Click to Login";

                            while (reader.Read())
                            {
                                Color MyColour = Color.PapayaWhip;
                                Color MyFontColour = Color.Black;
                                //MessageBox.Show(reader.GetValue(reader.GetOrdinal("UniID")).ToString());
                                if (reader.GetValue(reader.GetOrdinal("UniID")).ToString() == System.Security.Principal.WindowsIdentity.GetCurrent().Name)
                                {
                                    button6.BackColor = Color.YellowGreen;
                                    button6.Text = "logged in. Click to logout";
                                }

                                if (reader.GetValue(reader.GetOrdinal("UniID")).ToString().ToUpper() == "NORTHAMPTON\\BJBARRE")
                                {
                                    MyColour = Color.Blue;
                                    MyFontColour = Color.White;
                                }
                                if ((reader.GetValue(reader.GetOrdinal("UniID")).ToString()).ToUpper() == "NORTHAMPTON\\GETOWNS")
                                {
                                    MyColour = Color.Orchid;
                                    MyFontColour = Color.White;
                                }
                                if ((reader.GetValue(reader.GetOrdinal("UniID")).ToString()).ToUpper() == "NORTHAMPTON\\RJVADUK")
                                {
                                    MyColour = Color.Red;
                                    MyFontColour = Color.White;
                                }
                                if ((reader.GetValue(reader.GetOrdinal("UniID")).ToString()).ToUpper() == "NORTHAMPTON\\JGOUGH")
                                {
                                    MyColour = Color.Green;
                                    MyFontColour = Color.White;
                                }
                                if ((reader.GetValue(reader.GetOrdinal("UniID")).ToString()).ToUpper() == "NORTHAMPTON\\AHAMIL")
                                {
                                    MyColour = Color.Indigo;
                                    MyFontColour = Color.White;
                                }





                                if (buttonPhone1.BackColor == Color.Gainsboro)
                                {
                                    buttonPhone1.Tag = (reader.GetValue(reader.GetOrdinal("UniID")).ToString());
                                    buttonPhone1.Text = (reader.GetValue(reader.GetOrdinal("PhoneNumber")).ToString()).ToUpper();
                                    buttonPhone1.BackColor = MyColour;
                                    buttonPhone1.ForeColor = MyFontColour;

                                }

                                else if (buttonPhone2.BackColor == Color.Gainsboro)
                                {
                                    buttonPhone2.Tag = (reader.GetValue(reader.GetOrdinal("UniID")).ToString());
                                    buttonPhone2.Text = (reader.GetValue(reader.GetOrdinal("PhoneNumber")).ToString()).ToUpper();
                                    buttonPhone2.BackColor = MyColour;
                                    buttonPhone2.ForeColor = MyFontColour;
                                }
                                else if (buttonPhone3.BackColor == Color.Gainsboro)
                                {
                                    buttonPhone3.Tag = (reader.GetValue(reader.GetOrdinal("UniID")).ToString());
                                    buttonPhone3.Text = (reader.GetValue(reader.GetOrdinal("PhoneNumber")).ToString()).ToUpper();
                                    buttonPhone3.BackColor = MyColour;
                                    buttonPhone3.ForeColor = MyFontColour;
                                }
                                else if (buttonPhone4.BackColor == Color.Gainsboro)
                                {
                                    buttonPhone4.Tag = (reader.GetValue(reader.GetOrdinal("UniID")).ToString());
                                    buttonPhone4.Text = (reader.GetValue(reader.GetOrdinal("PhoneNumber")).ToString()).ToUpper();
                                    buttonPhone4.BackColor = MyColour;
                                    buttonPhone4.ForeColor = MyFontColour;
                                }
                                else if (buttonPhone5.BackColor == Color.Gainsboro)
                                {
                                    buttonPhone5.Tag = (reader.GetValue(reader.GetOrdinal("UniID")).ToString());
                                    buttonPhone5.Text = (reader.GetValue(reader.GetOrdinal("PhoneNumber")).ToString()).ToUpper();
                                    buttonPhone5.BackColor = MyColour;
                                    buttonPhone5.ForeColor = MyFontColour;
                                }
                                else if (buttonPhone6.BackColor == Color.Gainsboro)
                                {
                                    buttonPhone6.Tag = reader.GetValue(reader.GetOrdinal("UniID")).ToString();
                                    buttonPhone6.Text = (reader.GetValue(reader.GetOrdinal("PhoneNumber")).ToString()).ToUpper();
                                    buttonPhone6.BackColor = MyColour;
                                    buttonPhone6.ForeColor = MyFontColour;
                                }



                            }


                        }
                    }
                }



            }
            catch
            {

            }

        }

        public void CheckPhones()
        {
            /*
            if ((buttonPhone1.BackColor == Color.Gainsboro) && (notifyIcon.Visible == false))
            {
                notifyIcon.Visible = true;

                notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon.BalloonTipTitle = "UoN App Phone Alert";
                notifyIcon.BalloonTipText = "It appears that no one currently logged into the phones..." +
                                            Environment.NewLine +
                                            "";

                notifyIcon.ShowBalloonTip(5000);

            }
            else
            {
                notifyIcon.Visible = false;
            }
            //*/
        }



        void CheckReminder()
        {
            string currentoperater = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

            //MessageBox.Show(currentoperater);


            string s1 = currentoperater;
            string s2 = "getowns";
            bool b = s1.Contains(s2);

            //MessageBox.Show("" + currentoperater + " - " + b);

            if (b == true)
            {
                //MessageBox.Show("Exiting...");
                return;
            }
            //MessageBox.Show("Running...");

            if (DateTime.Now.Hour == 17)
            {

                //MessageBox.Show("" + DateTime.Now.ToString());

                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {

                        ArrayList UserList = new ArrayList();

                        Conn.Open();
                        //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)
                        //MessageBox.Show(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (NOT (LoanData.ID IN (SELECT Id FROM LoanData WHERE returned = collected AND returned not null))) AND CollectionDate > '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 01:00:00' and CollectionDate < '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 23:00:00'"; //AND LoanData.Collected = 0
                        cmd.CommandText = "Select distinct Customer, Firstname, Surname, Email, ReturnDate from LoanData LEFT JOIN Customer ON LoanData.Customer = Customer.ID where Returndate < '" + DateTime.Now.ToString("yyyy-MM-dd") + " 18:00:00" + "' and Returned isnull and Reminder1 isnull";


                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                //MessageBox.Show("" + reader.GetValue(reader.GetOrdinal("Customer")));
                                //MessageBox.Show("" + reader.GetValue(reader.GetOrdinal("Email")));
                                //MessageBox.Show("" + reader.GetValue(reader.GetOrdinal("ReturnDate")));

                                int UserID = Convert.ToInt32(reader.GetValue(reader.GetOrdinal("Customer")));
                                //MessageBox.Show("" + UserID);
                                UserList.Add(UserID);
                                SendReminder(Convert.ToString(reader.GetValue(reader.GetOrdinal("Email"))), Convert.ToString(reader.GetValue(reader.GetOrdinal("FirstName"))) + " " + Convert.ToString(reader.GetValue(reader.GetOrdinal("Surname"))), UserID);



                            }
                        }

                        foreach (int UserID in UserList)
                        {
                            String TimeNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            cmd.CommandText = "UPDATE LoanData SET Reminder1 = '" + TimeNow + "' where customer = " + UserID + " and reminder1 isnull";
                            cmd.ExecuteNonQuery();
                        }



                        Conn.Close();
                    }
                }

            }

        }


        private void SendReminder(String UserEmail, String FullName, Int32 UserID)
        {

            //MessageBox.Show("" + UserID);
            //UserEmail = "Simon.Ford@northampton.ac.uk";
            //UserEmail = "mark.rowland@northampton.ac.uk";
            //UserEmail = "James.Gough@northampton.ac.uk";
            //UserEmail = "INS_Service_Desk_Team@northampton.ac.uk";

            //MessageBox.Show("Token number is: " + System.Security.Principal.WindowsIdentity.GetCurrent().Token);

            /*
            MailMessage msg = new MailMessage();
            msg.To.Add(new MailAddress(UserEmail));
            msg.IsBodyHtml = true;
            msg.From = new MailAddress("AVBookings@northampton.ac.uk");
            msg.Subject = "Reminder to return your loaned equipment";
            //msg.Body = "<div>This is a HTML email test.</div>";
            */
            string FullEmail = @"<h1><span style=""color: #ff0000; font-family: OpenSans;"">" + FullName + ". Your loan equipment is now overdue&hellip;</span></h1>";

            FullEmail = FullEmail + File.ReadAllText(Directory.GetCurrentDirectory() + $"/Email-Reminder1-Start.html");

            string CreatedBy = System.Security.Principal.WindowsIdentity.GetCurrent().Name; //CreatedBy
            string LoanID = "";

            try
            {

                //Generate middle email from CusromerOrders Gridview.

                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                        //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)
                        //MessageBox.Show(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (NOT (LoanData.ID IN (SELECT Id FROM LoanData WHERE returned = collected AND returned not null))) AND CollectionDate > '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 01:00:00' and CollectionDate < '" + LoandateTimePicker.Value.ToString("yyyy-MM-dd") + " 23:00:00'"; //AND LoanData.Collected = 0
                        cmd.CommandText = "Select LoanDescriptions.Description, LoanStock.AssetID, Loandata.CollectionDate, LoanLocations.Location, LoanData.ReturnDate, LoanData.ID from LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID where Returndate < '" + DateTime.Now.ToString("yyyy-MM-dd") + " 17:00" + "' and Returned isnull and Customer = " + UserID;


                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {

                            while (reader.Read())
                            {
                                //MessageBox.Show("" + reader.GetValue(0));
                                FullEmail = FullEmail + "<tr><td>" + reader.GetValue(0) + "</td><td>" + reader.GetValue(1) + "</td><td>" + reader.GetValue(2) + "</td><td>" + reader.GetValue(3) + "</td><td>" + reader.GetValue(4) + "</td></tr>";
                                LoanID = reader.GetValue(5).ToString();
                            }
                        }



                        cmd.CommandText = "Update LoanData SET Reminder1 = '" + DateTime.Now + "', Reminder2 = '" + DateTime.Now.AddDays(7) + "',ReminderSentBy = '" + CreatedBy + "' Where LoanData.Id = " + LoanID;
                        cmd.ExecuteNonQuery();


                        Conn.Close();


                    }
                }
            }
            catch { }

            //Assemble full email.

            FullEmail = FullEmail + File.ReadAllText(Directory.GetCurrentDirectory() + $"/Email-Reminder1-End.html");

            // Add logo and contact info.

            LinkedResource LinkedImage = new LinkedResource(Environment.CurrentDirectory + @"\EmailLogo.png");
            LinkedImage.ContentId = "MyPic";

            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(FullEmail +
  @"<img src=cid:MyPic>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong><a href=""http://www.northampton.ac.uk/unit/"">Northampton.ac.uk</a></p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong  >University of Northampton,</strong> Grendon, Park Campus,</p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;"">Boughton Green Road, Northampton NN2 7AL United Kingdom</p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong>Follow the story on social media</strong></p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><a href=""http://www.northampton.ac.uk/social-media-hub/"">http://www.northampton.ac.uk/social-media-hub/</a></p>
                ",
  null, "text/html");


            Outlook.Application oApp = new Outlook.Application();
            Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            oMailItem.To = UserEmail;
            oMailItem.Subject = "Reminder to return your loaned equipment";
            oMailItem.HTMLBody = FullEmail;
            try { oMailItem.Display(false); } catch { }




            /*
            htmlView.LinkedResources.Add(LinkedImage);
            msg.AlternateViews.Add(htmlView);

            SmtpClient client = new SmtpClient();
            client.Host = "webmail.northampton.ac.uk";

            client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

            client.Port = 25;
            client.EnableSsl = false;
            client.UseDefaultCredentials = true;

            //MessageBox.Show("" + System.Net.CredentialCache.DefaultNetworkCredentials.UserName);

            //MessageBox.Show("Send Email");

            client.Send(msg);



            //*/
        }

        private void DatePlusbutton_Click(object sender, EventArgs e)
        {
            LoandateTimePicker.Value = LoandateTimePicker.Value.AddDays(1);
        }

        private void DateMinusbutton(object sender, EventArgs e)
        {
            LoandateTimePicker.Value = LoandateTimePicker.Value.AddDays(-1);
        }

        private void CollectsdataGridView_DoubleClick(object sender, EventArgs e)
        {

        }

        private void ReturnsdataGridView_DoubleClick(object sender, EventArgs e)
        {

        }


        private void LoandateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ReadTodaysBookings();
        }



        void WriteToLoanDB(String DBString)
        {
            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    Conn.Open();
                    //cmd.CommandText = "UPDATE LoanData SET Returned = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE Id = " + Convert.ToInt32(ReturnsdataGridView.CurrentRow.Cells[6].Value);
                    cmd.CommandText = DBString;
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "update Updates Set GetUpdates = 'last updated by " + System.Windows.Forms.SystemInformation.ComputerName + "' where hostname <> ''";
                    cmd.ExecuteNonQuery();

                    Conn.Close();
                }
            }
        }

        void WriteToPhoneDB(String DBString)
        {
            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\PhoneDB.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    Conn.Open();
                    //cmd.CommandText = "UPDATE LoanData SET Returned = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE Id = " + Convert.ToInt32(ReturnsdataGridView.CurrentRow.Cells[6].Value);
                    cmd.CommandText = DBString;
                    cmd.ExecuteNonQuery();

                    Conn.Close();



                }
            }

            string DBconString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            DBWright("update Updates Set GetUpdates = 'last updated by " + System.Windows.Forms.SystemInformation.ComputerName + "' where hostname <> ''", DBconString);

        }



        private void button1_Click(object sender, EventArgs e)
        {
            Form fc = Application.OpenForms["FormLoans"];

            if (fc != null)
                fc.Close();

            FormLoans Mainmenu = new FormLoans();
            Mainmenu.Show();
        }


        private void ManageLoans_Click(object sender, EventArgs e)
        {
            Form fc = Application.OpenForms["FormLoanManage"];

            if (fc != null)
                fc.Close();

            FormLoanManage Mainmenu = new FormLoanManage();
            Mainmenu.Show();
        }

        private void TabFlashtimer_Tick(object sender, EventArgs e)
        {
            if (currentColor == Color.Yellow)
            {
                currentColor = Color.Green;
                LoanstabControl.TabPages[2].Text = "ALERT! ATTENTION REQUIRED";
            }
            else
            {


                currentColor = Color.Yellow;
                LoanstabControl.TabPages[2].Text = "                                  ";
            }
            //LoanstabControl.Refresh();
        }

        private void LoanstabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            /*
            if (TabFlashtimer.Enabled && e.Index == 1)
            {
                e.Graphics.FillRectangle(new SolidBrush(currentColor), e.Bounds);
            }
            else
            {
                e.Graphics.FillRectangle(new SolidBrush(this.BackColor), e.Bounds);
            }
            Rectangle paddedBounds = e.Bounds;
            paddedBounds.Inflate(-2, -2);
            e.Graphics.DrawString(LoanstabControl.TabPages[e.Index].Text, this.Font, SystemBrushes.HighlightText, paddedBounds);
            //*/
        }

        private void CollectsdataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Convert.ToInt32(CollectsdataGridView.CurrentRow.Cells[6].Value) != 1)
            {
                //MessageBox.Show("" + System.Security.Principal.WindowsIdentity.GetCurrent().Name);

                String UserName = CollectsdataGridView.CurrentRow.Cells[0].Value.ToString();
                String Device = CollectsdataGridView.CurrentRow.Cells[1].Value.ToString();
                String AssetTag = CollectsdataGridView.CurrentRow.Cells[2].Value.ToString();
                Int32 UserID = Convert.ToInt32(CollectsdataGridView.CurrentRow.Cells[6].Value);

                String MSGMessage = "Has " + UserName + " collected the " + Device + " (" + AssetTag + ")?";
                String MSGTitle = "Confirm release";

                DialogResult dialogResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    WriteToLoanDB("UPDATE LoanData SET Collected = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', ReleasedBy = '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' WHERE Id = " + UserID);

                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }



            }

            ReadTodaysBookings();
        }

        private void ReturnsdataGridView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == false)
            {
                return;
            }


            if (Convert.ToInt32(ReturnsdataGridView.CurrentRow.Cells[6].Value) != 1)
            {
                //MessageBox.Show("" + System.Security.Principal.WindowsIdentity.GetCurrent().Name);

                String UserName = ReturnsdataGridView.CurrentRow.Cells[0].Value.ToString();
                String Device = ReturnsdataGridView.CurrentRow.Cells[1].Value.ToString();
                String AssetTag = ReturnsdataGridView.CurrentRow.Cells[2].Value.ToString();
                Int32 UserID = Convert.ToInt32(ReturnsdataGridView.CurrentRow.Cells[6].Value);

                String MSGMessage = "Has " + UserName + " returned the " + Device + " (" + AssetTag + ")?";
                String MSGTitle = "Confirm return";

                DialogResult dialogResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    WriteToLoanDB("UPDATE LoanData SET Returned = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', ReturnedBy = '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' WHERE Id = " + UserID);
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }

            }

            ReadTodaysBookings();
        }

        private void ReturnDevice_Click(object sender, EventArgs e)
        {

            String AssetTag = DevicesOutcomboBox.SelectedItem.ToString();
            String Device = "";
            String UserName = "";
            Int32 UserID = 0;
            Int32 LoanID = 0;
            String ReturnDate = "";

            if (DevicesOutcomboBox.SelectedItem != null)
            {
                //MessageBox.Show("" + System.Security.Principal.WindowsIdentity.GetCurrent().Name);

                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();


                        cmd.CommandText = @"
                                        Select  LoanData.ID, FirstName, Surname,ReturnDate from LoanData 
                                        left join Loanstock on Loandata.Device = LoanStock.ID
                                        left join Customer on LoanData.Customer = Customer.ID
                                        where Collected is not null and returned is null and AssetID = '" + AssetTag + @"'
                                        order by AssetID Asc
                                        ";

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {

                                UserName = reader.GetValue(reader.GetOrdinal("FirstName")).ToString() + " " + reader.GetValue(reader.GetOrdinal("Surname")).ToString();
                                LoanID = Convert.ToInt32(reader.GetValue(reader.GetOrdinal("ID")));
                                ReturnDate = reader.GetValue(reader.GetOrdinal("ReturnDate")).ToString();


                            }
                        }




                        String MSGMessage = AssetTag + " is currently loaned out to " + UserName + " and is due back on " + ReturnDate + ". Do you want to return this now?";
                        String MSGTitle = "Confirm return";

                        DialogResult dialogResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            WriteToLoanDB("UPDATE LoanData SET Returned = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', ReturnedBy = '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' WHERE Id = " + LoanID);
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                            //do something else
                        }

                        Conn.Close();

                    }
                    DevicesOutcomboBox.Text = "";
                    ReadTodaysBookings();
                }
            }

        }

        private void labelManage_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            Form fc = Application.OpenForms["FormLoanManage"];

            if (fc != null)
                fc.Close();

            FormLoanManage Mainmenu = new FormLoanManage();
            Mainmenu.Show();
            //*/
        }

        private void buttonCMC_Click(object sender, EventArgs e)
        {
            //C:\Windows\System32\runas.exe /user:northampton\MDRowla_Admin /savecred  "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe"
            /*
            Process procAD = new Process();
            procAD.StartInfo.UseShellExecute = false;
            //procAD.StartInfo.Verb = "runas";
            procAD.StartInfo.FileName = "C:\\Program Files (x86)\\Microsoft Configuration Manager\\AdminConsole\\bin\\Microsoft.ConfigurationManagement.exe";
            procAD.StartInfo.UserName = "NORTHAMPTON\\MDRowla_Admin";
        
            //procAD.StartInfo.Arguments = "/user:northampton\\MDRowla_Admin /savecred ''";
            procAD.Start();
            //*/


        }

        private void buttonAD_Click(object sender, EventArgs e)
        {


        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (button6.BackColor == Color.Goldenrod)
            {
                button6.BackColor = Color.YellowGreen;
                //MessageBox.Show("Here");
                button6.Text = "logged in. Click to logout";
                Thread t = new Thread(() =>
                {
                    WriteToPhoneDB("INSERT INTO Activeusers(UniID,Login,Checkin, PhoneNumber) VALUES ('" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + textBoxPhoneNumber.Text + "')");

                }
                );
                t.Start();

            }
            else if (button6.BackColor == Color.YellowGreen)
            {
                //MessageBox.Show("Here");
                button6.BackColor = Color.Goldenrod;
                button6.Text = "Click to Login";
                Thread t = new Thread(() =>
                {
                    WriteToPhoneDB("UPDATE Activeusers SET Logout = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE UniID = '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' AND Logout is null");


                }
                );
                t.Start();

            }


        }

        void buttonPhone1_MouseHover(object sender, EventArgs e)
        {
            if (buttonPhone1.BackColor != Color.Gainsboro)
                toolTipPhone1.Show("" + buttonPhone1.Tag.ToString(), buttonPhone1);
        }
        private void buttonPhone2_MouseHover(object sender, EventArgs e)
        {
            if (buttonPhone2.BackColor != Color.Gainsboro)
                toolTipPhone2.Show("" + buttonPhone2.Tag.ToString(), buttonPhone2);
        }

        private void buttonPhone3_MouseHover(object sender, EventArgs e)
        {
            if (buttonPhone3.BackColor != Color.Gainsboro)
                toolTipPhone3.Show("" + buttonPhone3.Tag.ToString(), buttonPhone3);
        }

        private void buttonPhone4_MouseHover(object sender, EventArgs e)
        {
            if (buttonPhone4.BackColor != Color.Gainsboro)
                toolTipPhone4.Show("" + buttonPhone4.Tag.ToString(), buttonPhone4);
        }

        private void buttonPhone5_MouseHover(object sender, EventArgs e)
        {
            if (buttonPhone5.BackColor != Color.Gainsboro)
                toolTipPhone5.Show("" + buttonPhone5.Tag.ToString(), buttonPhone5);
        }

        private void buttonPhone6_MouseHover(object sender, EventArgs e)
        {
            if (buttonPhone6.BackColor == Color.LawnGreen)
                toolTipPhone6.Show("" + buttonPhone6.Tag.ToString(), buttonPhone6);
        }



        public void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                //MessageBox.Show("Locked");
                // Add your session lock "handling" code here
                button6.BackColor = Color.Goldenrod;
                button6.Text = "Click to Login";
                WriteToPhoneDB("UPDATE Activeusers SET Logout = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE UniID = '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' AND Logout is null");

            }

        }
        private void forceLogOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Here");
            //MessageBox.Show(forceLogOutToolStripMenuItem.AccessibleDescription);
            WriteToPhoneDB("UPDATE Activeusers SET Logout = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE UniID = '" + forceLogOutToolStripMenuItem.AccessibleDescription + "' AND Logout is null");

        }
        private void buttonPhone1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                forceLogOutToolStripMenuItem.AccessibleDescription = "";
                forceLogOutToolStripMenuItem.Text = "No Options Here";
                if (e.Button == MouseButtons.Right && buttonPhone1.BackColor != Color.Gainsboro)
                {
                    forceLogOutToolStripMenuItem.AccessibleDescription = buttonPhone1.Tag.ToString();
                    forceLogOutToolStripMenuItem.Text = "Log out " + buttonPhone1.Tag.ToString();
                }
            }
            catch { }
        }

        private void buttonPhone2_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                forceLogOutToolStripMenuItem.AccessibleDescription = "";
                forceLogOutToolStripMenuItem.Text = "No Options Here";
                if (e.Button == MouseButtons.Right && buttonPhone2.BackColor != Color.Gainsboro)
                {
                    forceLogOutToolStripMenuItem.AccessibleDescription = buttonPhone2.Tag.ToString();
                    forceLogOutToolStripMenuItem.Text = "Log out " + buttonPhone2.Tag.ToString();
                }
            }
            catch { }
        }

        private void buttonPhone3_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                forceLogOutToolStripMenuItem.AccessibleDescription = "";
                forceLogOutToolStripMenuItem.Text = "No Options Here";
                if (e.Button == MouseButtons.Right && buttonPhone3.BackColor != Color.Gainsboro)
                {
                    forceLogOutToolStripMenuItem.AccessibleDescription = buttonPhone3.Tag.ToString();
                    forceLogOutToolStripMenuItem.Text = "Log out " + buttonPhone3.Tag.ToString();
                }
            }
            catch { }
        }

        private void buttonPhone4_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                forceLogOutToolStripMenuItem.AccessibleDescription = "";
                forceLogOutToolStripMenuItem.Text = "No Options Here";
                if (e.Button == MouseButtons.Right && buttonPhone4.BackColor != Color.Gainsboro)
                {
                    forceLogOutToolStripMenuItem.AccessibleDescription = buttonPhone4.Tag.ToString();
                    forceLogOutToolStripMenuItem.Text = "Log out " + buttonPhone4.Tag.ToString();
                }
            }
            catch { }
        }

        private void buttonPhone5_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                forceLogOutToolStripMenuItem.AccessibleDescription = "";
                forceLogOutToolStripMenuItem.Text = "No Options Here";
                if (e.Button == MouseButtons.Right && buttonPhone5.BackColor != Color.Gainsboro)
                {
                    forceLogOutToolStripMenuItem.AccessibleDescription = buttonPhone5.Tag.ToString();
                    forceLogOutToolStripMenuItem.Text = "Log out " + buttonPhone5.Tag.ToString();
                }
            }
            catch { }
        }

        private void buttonPhone6_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                forceLogOutToolStripMenuItem.AccessibleDescription = "";
                if (e.Button == MouseButtons.Right && buttonPhone6.BackColor != Color.Gainsboro)
                {
                    forceLogOutToolStripMenuItem.AccessibleDescription = buttonPhone6.Tag.ToString();
                    forceLogOutToolStripMenuItem.Text = "Log out " + buttonPhone6.Tag.ToString();
                }
            }
            catch { }
        }

        private void contextMenuStrip2_Opened(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (SEANSSave == true)
            {
                SEANSSave = false;
                pictureBox1.Image = Properties.Resources.ToggleOff;
                pictureBox1.Refresh();

            }
            else
            {
                SEANSSave = true;
                pictureBox1.Image = Properties.Resources.ToggleOn;
                pictureBox1.Refresh();
            }



        }





        private void textBoxE_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabControl2_Leave(object sender, EventArgs e)
        {

        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (tabControl1.SelectedIndex == 2)
            {
                timerChase.Enabled = true;
                ChaseLoad();

                dataGridViewChaseMain.Font = new Font("Open Sans", 9F, GraphicsUnit.Pixel);
                dataGridViewChaseDetail.Font = new Font("Open Sans", 9F, GraphicsUnit.Pixel);

                for (int i = 0; i <= dataGridViewChaseMain.ColumnCount - 1; i++)
                {

                    if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "TicketID") { ChaseTicketIDCol = i; }
                    else if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "UserEmail") { ChaseEmailIDCol = i; }
                    else if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "Status") { ChaseStatusIDCol = i; }
                    else if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "InfoRequest") { ChaseInfoIDCol = i; }
                    else if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "Reference") { ChaseTicketRef = i; }
                    else if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "TicketTitle") { ChaseTicketTitle = i; }
                    else if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "UsersName") { ChaseUserCol = i; }
                    else if (dataGridViewChaseMain.Columns[i].HeaderText.ToString() == "NextReminderDate") { ChaseReminderCol = i; }





                }
            }
            else
            {
                timerChase.Enabled = false;
            }
        }

        private void timerChase_Tick(object sender, EventArgs e)
        {
            ChaseLoad();

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            Chase_Add_new();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Chase_Resolve();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Chase_SendEmail();
        }

        public void ChaseLoad()
        {
            //Ticket Chase
            int CurrentRow = 0;
            int CurrentCell = 0;
            int ScrollRow = 0;
            int ScrollCol = 0;

            int LoanCurrentRow = 0;
            int LoanCurrentCell = 0;
            int LoanScrollRow = 0;
            int LoanScrollCol = 0;


            //try {MessageBox.Show(dataGridViewChaseMain.FirstDisplayedScrollingRowIndex.ToString());} catch { }

            try { CurrentRow = dataGridViewChaseMain.CurrentCell.RowIndex; } catch { }
            try { CurrentCell = dataGridViewChaseMain.CurrentCell.ColumnIndex; } catch { }
            try { ScrollRow = dataGridViewChaseMain.FirstDisplayedScrollingRowIndex; } catch { }
            try { ScrollCol = dataGridViewChaseMain.FirstDisplayedScrollingColumnIndex; } catch { }

            //dataGridViewChaseMain.Columns.Clear();
            dataGridViewChaseMain.AutoGenerateColumns = true;
            dataGridViewChaseMain.DataSource = null;
            //ataGridViewChaseMain.Rows.Clear();

            string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\ChaseDB.db3";
            string SQLString = "Select * from Ticket Where CurrentOwner = '" + OperaterID + "' and (Status = 0 or Status = 1 or Status = 2 or Status = 3)";

            dataGridViewChaseMain.DataSource = GetDataSet(conString, SQLString);


            try { dataGridViewChaseMain.CurrentCell = this.dataGridViewChaseMain[CurrentCell, CurrentRow]; } catch { }
            try { dataGridViewChaseMain.FirstDisplayedScrollingRowIndex = ScrollRow; } catch { }
            try { dataGridViewChaseMain.FirstDisplayedScrollingColumnIndex = ScrollCol; } catch { }



            try
            {


                foreach (DataGridViewRow row in dataGridViewChaseMain.Rows)

                    if (Convert.ToDateTime(row.Cells[ChaseReminderCol].Value.ToString()) < DateTime.Now)
                    {
                        row.DefaultCellStyle.BackColor = Color.Orange;
                    }

            }
            catch { }
            //Loan Chase

            try { LoanCurrentRow = dataGridViewChaseLoan.CurrentCell.RowIndex; } catch { }
            try { LoanCurrentCell = dataGridViewChaseLoan.CurrentCell.ColumnIndex; } catch { }
            try { LoanScrollRow = dataGridViewChaseLoan.FirstDisplayedScrollingRowIndex; } catch { }
            try { LoanScrollCol = dataGridViewChaseLoan.FirstDisplayedScrollingColumnIndex; } catch { }

            dataGridViewChaseLoan.Font = new Font("Open Sans", 9F, GraphicsUnit.Pixel);
            dataGridViewChaseLoanDetail.Font = new Font("Open Sans", 9F, GraphicsUnit.Pixel);

            string Today = DateTime.Now.ToString("yyyy-MM-dd") + " 10:00:00";

            dataGridViewChaseLoan.Columns.Clear();
            dataGridViewChaseLoan.AutoGenerateColumns = true;
            dataGridViewChaseLoan.DataSource = null;
            //dataGridViewChaseLoan.Rows.Clear();

            conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            SQLString = @"
            select LoanData.ID as LID, Customer.ID as CID, Customer.FirstName, Customer.Surname, Customer.Telephone, LoanStock.AssetID, strftime('%d-%m-%Y',LoanData.ReturnDate) as 'Agreed Return Date', LoanData.Reminder1, LoanData.Reminder2 as 'Next Reminder Date', LoanData.ReminderSentBy, Customer.Email from LoanData
            left join Customer on Loandata.Customer = Customer.ID
            left join LoanStock on LoanData.Device = LoanStock.ID
            where LoanData.Returned is null and ReturnDate <'" + Today + "'";


            dataGridViewChaseLoan.DataSource = GetDataSet(conString, SQLString);
            try
            {


                foreach (DataGridViewRow row in dataGridViewChaseLoan.Rows)

                    if (Convert.ToDateTime(row.Cells[FindColID(dataGridViewChaseLoan, "Next Reminder Date")].Value.ToString()) < DateTime.Now)
                    {
                        row.DefaultCellStyle.BackColor = Color.Orange;
                    }

            }
            catch { }


            try { dataGridViewChaseLoan.CurrentCell = this.dataGridViewChaseLoan[LoanCurrentCell, LoanCurrentRow]; } catch { }
            try { dataGridViewChaseLoan.FirstDisplayedScrollingRowIndex = LoanScrollRow; } catch { }
            try { dataGridViewChaseLoan.FirstDisplayedScrollingColumnIndex = LoanScrollCol; } catch { }

        }

        private void dataGridViewChaseMain_SelectionChanged(object sender, EventArgs e)
        {

        }

        public DataTable GetDataSet(string ConnectionString, string SQL)
        {
            DataTable dt = new DataTable();

            try
            {
                using (var c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(SQL, c))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            dt.Load(rdr);
                        }
                    }
                }
            }
            catch { }
            return dt;
        }

        public void DBWright(string SQLString, string Connection)
        {
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection(Connection))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    Conn.Open();

                    cmd.CommandText = SQLString;
                    cmd.ExecuteNonQuery();

                    Conn.Close();
                }
            }
        }

        static string DBSingleRead(string SQLString, string Connection)
        {
            String DBValue = "";

            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection(Connection))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    try
                    {
                        Conn.Open();
                        cmd.CommandText = SQLString;
                        DBValue = cmd.ExecuteScalar().ToString();
                        Conn.Close();
                    }
                    catch { }
                }
            }



            return DBValue;
        }

        static int ChaseCountDB(string SQLString, string conString)
        {
            Int32 count = 0;

            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection(conString))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    try
                    {
                        Conn.Open();
                        cmd.CommandText = SQLString;
                        count = Convert.ToInt32(cmd.ExecuteScalar());
                        Conn.Close();
                        //MessageBox.Show("" + count);
                    }
                    catch { }
                }
            }

            return count;
        }

        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }

        public void Chase_Add_new()
        {
            try
            {
                //textBox1.Height = 300;
                //textBox1.Text = "";
                //textBox1.Text = Clipboard.GetText();

                string ClipboardText = Clipboard.GetText().Replace("'", "");

                string TicketRef = getBetween(ClipboardText, "Help", "Major Incident");
                TicketRef = TicketRef.Substring(TicketRef.Length - 17);
                TicketRef = TicketRef.Replace(" ", "").Replace("\n", "").Replace("\r", "");
                //MessageBox.Show(TicketRef);

                string Contactx = getBetween(ClipboardText, "Contact", "Telephone");
                Contactx = Contactx.Replace(" ", "").Replace("\n", "").Replace("\r", "");
                Contactx = "STARTSTART" + Contactx + "ENDEND";
                string Contact = getBetween(Contactx, ",", "ENDEND") + " " + getBetween(Contactx, "STARTSTART", ",");
                //MessageBox.Show(Contact);

                //Description
                string Description = getBetween(ClipboardText, "Description", "Allocation");
                Description = Description.Replace("\n", "").Replace("\r", "");
                //MessageBox.Show(Description);

                string Email = getBetween(ClipboardText, "Using", "Third Party");
                Email = Email.Replace(" ", "").Replace("\n", "").Replace("\r", "");
                //MessageBox.Show(Email);

                string Notes = getBetween(ClipboardText, "DatePrivatePublicAudit", "SubmitSubmit");

                //Load All Existing Notes...
                Notes = Notes + "ENDEND";

                /* Convert the string into a byte[].
                byte[] asciiBytes = Encoding.ASCII.GetBytes(Notes);
                string result = string.Join(",", asciiBytes);
                MessageBox.Show(result);
                */

                //Notes = getBetween(Notes, "Public", "ENDEND");
                Notes = getBetween(Notes, "\r", "ENDEND");
                Notes = Notes + "Public";
                Notes = getBetween(Notes, "\r", "Public");
                Notes = Notes.Replace("Load All Existing Notes...", "");
                //MessageBox.Show(Notes);


                string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\ChaseDB.db3";
                string SQLString = "select count (Reference) from Ticket where Reference = '" + TicketRef + "'";

                //MessageBox.Show(ChaseCountDB(SQLString, conString).ToString());

                if (ChaseCountDB(SQLString, conString) == 0)
                {
                    try
                    {
                        //Ticket
                        SQLString = @"insert into ticket 
(Reference, CreatedBy, CreatedDate, CurrentOwner, Status, UsersName, Notes, InfoRequest, UserEmail, TicketTitle, NextReminderDate) 
Values ('" + TicketRef + "','" + OperaterID + "','" + DateTime.Now + "','" + OperaterID + "',0,'" + Contact + "','','" + Notes + "','" + Email + "', '" + Description + "', '" + DateTime.Now + "')";

                        DBWright(SQLString, conString);

                        //Get New ID
                        SQLString = "select TicketID from Ticket where Reference = '" + TicketRef + "'";
                        int NewID = ChaseCountDB(SQLString, conString);
                        //MessageBox.Show("" + NewID);
                        //Ledger
                        SQLString = "insert into Ledger (TicketID, User, Date, Action, Notes) Values (" + NewID + ",'" + OperaterID + "','" + DateTime.Now + "','Added to database','')";

                        DBWright(SQLString, conString);
                    }
                    catch { }
                }
                else
                {

                    SQLString = "update Ticket Set CurrentOwner = '" + OperaterID + "' where Reference = '" + TicketRef + "'";
                    DBWright(SQLString, conString);

                    SQLString = "select TicketID from Ticket where Reference = '" + TicketRef + "'";
                    int NewID = ChaseCountDB(SQLString, conString);

                    SQLString = "insert into Ledger (TicketID, User, Date, Action, Notes) Values (" + NewID + ",'" + OperaterID + "','" + DateTime.Now + "','Changed Owner','')";
                    //MessageBox.Show("here");
                    DBWright(SQLString, conString);

                    SQLString = "select Status from Ticket where Reference = '" + TicketRef + "'";

                    string TicketStatus = DBSingleRead(SQLString, conString);

                    //MessageBox.Show(TicketStatus);

                    if (TicketStatus == "4")
                    {
                        MessageBox.Show("Already exceeded stage 3");
                        SQLString = "update Ticket Set Status = 2 where Reference = '" + TicketRef + "'";
                        DBWright(SQLString, conString);
                        SQLString = "insert into Ledger (TicketID, User, Date, Action, Notes) Values (" + NewID + ",'" + OperaterID + "','" + DateTime.Now + "','Re-activate - Status changed to 2','')";
                        DBWright(SQLString, conString);

                    }



                }
                ChaseLoad();
            }
            catch
            {
                string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\ChaseDB.db3";
                string SQLString = "insert into ClipboardErrors (ClipboardValue, OperaterID, Date) Values ('" + Clipboard.GetText().Replace("'", "") + "','" + OperaterID + "','" + DateTime.Now + "')";

                DBWright(SQLString, conString);
                MessageBox.Show("There was a problem capturing the information. Please copy it from Marvel and try again");
                return;
            }
        }

        public void Chase_Resolve()
        {
            if (dataGridViewChaseMain.CurrentCell.RowIndex == -1) { return; }
            try
            {


                string TicketID = (dataGridViewChaseMain.CurrentRow.Cells[ChaseTicketIDCol].Value.ToString());


                string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\ChaseDB.db3";
                string SQLString = "update Ticket Set Status = 4 where TicketID = " + TicketID;
                DBWright(SQLString, conString);

                SQLString = "insert into Ledger (TicketID, User, Date, Action, Notes) Values (" + TicketID + ",'" + OperaterID + "','" + DateTime.Now + "','Maked as resolved','')";
                DBWright(SQLString, conString);
                ChaseLoad();
            }
            catch { }




            //    gridView1.HeaderRow.Cells[0].Text
            //dataGridViewChaseMain.col


        }

        public void Chase_SendEmail()
        {
            try { if (dataGridViewChaseMain.CurrentCell.RowIndex == -1) { return; } } catch { return; }
            if (dataGridViewChaseMain.ColumnCount < 3) { return; }

            if (Convert.ToDateTime(dataGridViewChaseMain.CurrentRow.Cells[ChaseReminderCol].Value.ToString()) > DateTime.Now) { return; }

            int CurrentRow = dataGridViewChaseMain.CurrentCell.RowIndex;
            string Status = "";



            Outlook.Application oApp = new Outlook.Application();
            Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

            Outlook.Attachments attachments = null;
            Outlook.Attachment attachment = null;











            string address = (dataGridViewChaseMain.CurrentRow.Cells[ChaseEmailIDCol].Value.ToString());
            string user = (dataGridViewChaseMain.CurrentRow.Cells[ChaseUserCol].Value.ToString());
            string inforrequest = (dataGridViewChaseMain.CurrentRow.Cells[ChaseInfoIDCol].Value.ToString());
            string TicketRef = (dataGridViewChaseMain.CurrentRow.Cells[ChaseTicketRef].Value.ToString());
            string TicketTitle = (dataGridViewChaseMain.CurrentRow.Cells[ChaseTicketTitle].Value.ToString());
            string TicketID = (dataGridViewChaseMain.CurrentRow.Cells[ChaseTicketIDCol].Value.ToString());
            int TicketStatus = Int32.Parse(dataGridViewChaseMain.CurrentRow.Cells[ChaseStatusIDCol].Value.ToString());

            if (dataGridViewChaseMain.Rows[CurrentRow].Cells[ChaseStatusIDCol].Value.ToString() == "0") { Status = "Thank you for your service request: "; }
            if (dataGridViewChaseMain.Rows[CurrentRow].Cells[ChaseStatusIDCol].Value.ToString() == "1") { Status = "This is our second attempt to contact you regarding your request: "; }
            if (dataGridViewChaseMain.Rows[CurrentRow].Cells[ChaseStatusIDCol].Value.ToString() == "2") { Status = "This is our third and final attempt to contact you regarding your request: "; }
            if (dataGridViewChaseMain.Rows[CurrentRow].Cells[ChaseStatusIDCol].Value.ToString() == "3") { return; }





            string body = @"<body><p><h1 style=""text - align: center; font - family: OpenSans; color: #5e9ca0;"">Further information is required to process your request with IT Services.</h1></p>

      <p style=""color: #2e6c80; font-family: OpenSans;"" ><font size='4'><b> Dear " + user + ",</font></b></p>";


            body = body + "<p>" + Status + " <b>" + TicketTitle + @".</b></p>

            <p>To process your ticket further information has been requested: <b>" + inforrequest + @"</b></p>

            <p>Your ticket will now be placed on hold until we have received the required information either by replying to this email or contacting IT Services on 01604 89 (3333) quoting your reference number <b>" + TicketRef + @"</b>.</p><p></p>

            <p>Please be aware that we operate a three contact process where we will email you on three separate occasions over a period of five days. After which time if we have not received a response your ticket will be resolved and you may be required to submit a new request.</p><p></p>

";


            body = body + @"            <p>Thank you.</p>

 <p> IT Services </p>
<p> 01604 893333 </p> ";

            string imageFile = Environment.CurrentDirectory + @"\EmailLogo.png";
            string imageCid = "image001.UoNLogo.png";


            attachments = oMailItem.Attachments;
            attachment = attachments.Add(imageFile,
               Outlook.OlAttachmentType.olEmbeddeditem, null, "");

            attachment.PropertyAccessor.SetProperty(
              "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
             , imageCid
             );

            body = body + String.Format(
               "<p><img src=\"cid:{0}\"></p>"
             , imageCid
             );


            body = body + @"<p style = ""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong><a href=""http://www.northampton.ac.uk/unit/"">Northampton.ac.uk</a></p>
            <p style = ""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong  >University of Northampton,</strong> Grendon, Park Campus,</p>
            <p style = ""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;"">Boughton Green Road, Northampton NN2 7AL United Kingdom</p>
            <p style = ""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
            <p style = ""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong>Follow the story on social media</strong></p>
            <p style = ""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><a href=""http://www.northampton.ac.uk/social-media-hub/"">http://www.northampton.ac.uk/social-media-hub/</a></p>
                
            ";


            //MessageBox.Show("Here");
            oMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            oMailItem.HTMLBody = body;
            oMailItem.To = address;
            oMailItem.Subject = "UoN IT Services: Information Request. Ref: " + TicketRef;

            oMailItem.Display(false);

            string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\ChaseDB.db3";
            string SQLString = "update ticket set Status = " + (TicketStatus + 1) + " where TicketId = " + TicketID;
            DBWright(SQLString, conString);

            //MessageBox.Show("" + (DateTime.Now.AddDays(2).DayOfWeek));

            if ((TicketStatus + 1 == 1) || (TicketStatus + 1 == 2))
            {
                if (DateTime.Now.AddDays(2).DayOfWeek == DayOfWeek.Saturday) { SQLString = "update ticket set NextReminderDate = '" + DateTime.Now.AddDays(4) + "' where TicketId = " + TicketID; }
                else if (DateTime.Now.AddDays(2).DayOfWeek == DayOfWeek.Sunday) { SQLString = "update ticket set NextReminderDate = '" + DateTime.Now.AddDays(3) + "' where TicketId = " + TicketID; }
                else { SQLString = "update ticket set NextReminderDate = '" + DateTime.Now.AddDays(2) + "' where TicketId = " + TicketID; }
                DBWright(SQLString, conString);
            }
            if (TicketStatus + 1 == 3)
            {
                if (DateTime.Now.AddDays(2).DayOfWeek == DayOfWeek.Saturday) { SQLString = "update ticket set NextReminderDate = '" + DateTime.Now.AddDays(3) + "' where TicketId = " + TicketID; }
                else if (DateTime.Now.AddDays(2).DayOfWeek == DayOfWeek.Sunday) { SQLString = "update ticket set NextReminderDate = '" + DateTime.Now.AddDays(2) + "' where TicketId = " + TicketID; }
                else { SQLString = "update ticket set NextReminderDate = '" + DateTime.Now.AddDays(1) + "' where TicketId = " + TicketID; }
                DBWright(SQLString, conString);
            }




            SQLString = "insert into Ledger (TicketID, User, Date, Action, Notes) Values (" + TicketID + ",'" + OperaterID + "','" + DateTime.Now + "','Sent confirmation (" + (TicketStatus + 1) + " of 3)','')";
            DBWright(SQLString, conString);

            var TextString = ("S (Situation):More information needed" + Environment.NewLine +
          "E (Escalation):" + Environment.NewLine +
          "A (Action):Sent user an email (" + (TicketStatus + 1) + " of 3)" + Environment.NewLine +
          "N (Next Step):Wait for reply" + Environment.NewLine +
          "S (Solution):" + Environment.NewLine +
          Environment.NewLine + "Notes: ");

            Clipboard.SetText(TextString);


            if (TicketID == "3")
            {
                int TicketIDNumber = Int32.Parse(TicketRef.Substring(TicketRef.Length - 5)) + 162;
                Process.Start("chrome.exe", "https://itservicedesk.northampton.ac.uk/MSM/RFP/Forms/Request.aspx?id=" + TicketIDNumber);
            }



        }

        private void dataGridViewChaseMain_DoubleClick(object sender, EventArgs e)
        {
            string TicketRef = (dataGridViewChaseMain.CurrentRow.Cells[ChaseTicketRef].Value.ToString());
            //MessageBox.Show(TicketRef);
            int TicketIDNumber = Int32.Parse(TicketRef.Substring(TicketRef.Length - 5)) + 162;

            Process.Start("chrome.exe", "https://itservicedesk.northampton.ac.uk/MSM/RFP/Forms/Request.aspx?id=" + TicketIDNumber);
        }

        void ChaseLoadDetails()
        {
            try
            {
                string TicketID = (dataGridViewChaseMain.CurrentRow.Cells[ChaseTicketIDCol].Value.ToString());

                string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\ChaseDB.db3";
                string SQLString = "Select Action, Date, User from ledger Where TicketID = " + TicketID + " order by Date Desc";

                DataTable dt = GetDataSet(conString, SQLString);
                dataGridViewChaseDetail.DataSource = dt;
            }
            catch { }
        }

        private void textBoxS_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ReadData();
            }
            catch { }
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {



        }

        private void hScrollBar1_SizeChanged(object sender, EventArgs e)
        {

        }

        private void hScrollBar1_ValueChanged(object sender, EventArgs e)
        {
            Screen screen = Screen.PrimaryScreen;
            int S_width = screen.Bounds.Width;
            int S_height = screen.Bounds.Height;

            int NewWidth = (375 + (hScrollBar1.Value * 5));

            this.MaximumSize = new System.Drawing.Size(NewWidth, S_height);
            this.Width = NewWidth;

            Properties.Settings.Default.FormWidth = hScrollBar1.Value;
            Properties.Settings.Default.Save();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void dataGridViewChaseMain_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            ChaseLoadDetails();
        }

        private void dataGridViewChaseLoan_DoubleClick(object sender, EventArgs e)
        {


        }

        private void buttonLoanSendReminder_Click(object sender, EventArgs e)
        {
            string UserEmail = "";
            string FullName = "";
            int UserID = 1;

            UserEmail = dataGridViewChaseLoan.CurrentRow.Cells[FindColID(dataGridViewChaseLoan, "Email")].Value.ToString();
            FullName = dataGridViewChaseLoan.CurrentRow.Cells[FindColID(dataGridViewChaseLoan, "FirstName")].Value.ToString() + " " + dataGridViewChaseLoan.CurrentRow.Cells[FindColID(dataGridViewChaseLoan, "Surname")].Value.ToString();
            UserID = Convert.ToInt32(dataGridViewChaseLoan.CurrentRow.Cells[FindColID(dataGridViewChaseLoan, "CID")].Value.ToString());

            SendReminder(UserEmail, FullName, UserID);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBoxPhoneNumber_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PhoneNumber = textBoxPhoneNumber.Text;
            Properties.Settings.Default.Save();
        }

        private void buttonRC_Click(object sender, EventArgs e)
        {
            //Process.Start("ConsoleApplication2.exe");
        }

        private void buttonRA_Click(object sender, EventArgs e)
        {


        }


        private void notifyIcon_DoubleClick(object Sender, EventArgs e)
        {
            // Show the form when the user double clicks on the notify icon.

            // Set the WindowState to normal if the form is minimized.
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal;

            // Activate the form.
            this.Activate();
        }

        private void menuItem_Click(object Sender, EventArgs e)
        {
            // Close the form, which closes the application.
            this.Close();
        }

        private void timerPhoneAlert_Tick(object sender, EventArgs e)
        {
            /*
            if (notifyIcon.Visible == false)
            {
                timerPhoneAlert.Interval = 250;
                notifyIcon.Visible = true;
            }
            else
            {
                
                notifyIcon.Visible = false;
            }
            */
        }

        private void tabControl2_Enter(object sender, EventArgs e)
        {
            string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            string SQLString = "";

            SQLString = "select * from LoanLocations";
            DGV_DHCampus.DataSource = GetDataSet(conString, SQLString);
            for (int i = 0; i <= DGV_DHCampus.ColumnCount - 1; i++)
            {
                if (DGV_DHCampus.Columns[i].HeaderText.ToString() == "ID") { DGV_DHCampus.Columns[i].Visible = false; }
            }


            SQLString = "select * from LoanDescriptions";
            DGV_DHType.DataSource = GetDataSet(conString, SQLString);
            for (int i = 0; i <= DGV_DHType.ColumnCount - 1; i++)
            {
                if (DGV_DHType.Columns[i].HeaderText.ToString() == "ID") { DGV_DHType.Columns[i].Visible = false; }
                if (DGV_DHType.Columns[i].HeaderText.ToString() == "SortOrder") { DGV_DHType.Columns[i].Visible = false; }
            }
        }


        private void Load_DHData()
        {
            DGV_DHHistory.DataSource = null;
            string LocID = DGV_DHCampus.CurrentRow.Cells[FindColID(DGV_DHCampus, "ID")].Value.ToString();
            string TypeID = DGV_DHType.CurrentRow.Cells[FindColID(DGV_DHType, "ID")].Value.ToString();

            string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            string SQLString = "select * from LoanStock Where Location = " + LocID + " and Description = " + TypeID;

            DGV_DHDevice.DataSource = GetDataSet(conString, SQLString);

            for (int i = 0; i <= DGV_DHDevice.ColumnCount - 1; i++)
            {
                if (DGV_DHDevice.Columns[i].HeaderText.ToString() == "ID") { DGV_DHDevice.Columns[i].Visible = false; }
                if (DGV_DHDevice.Columns[i].HeaderText.ToString() == "Description") { DGV_DHDevice.Columns[i].Visible = false; }
                if (DGV_DHDevice.Columns[i].HeaderText.ToString() == "Location") { DGV_DHDevice.Columns[i].Visible = false; }
                if (DGV_DHDevice.Columns[i].HeaderText.ToString() == "Active") { DGV_DHDevice.Columns[i].Visible = false; }
            }


        }

        private void DGV_DHType_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Load_DHData();
        }

        private void DGV_DHCampus_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Load_DHData();
        }

        private void DGV_DHDevice_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string DevID = DGV_DHDevice.CurrentRow.Cells[FindColID(DGV_DHDevice, "ID")].Value.ToString();
            string conString = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            string SQLString = @"select FirstName, Surname, CreatedBy, CollectionDate, Collected, ReleasedBy, ReturnDate, Returned, ReturnedBy, Telephone, Mobile  from loandata 
                                left join Customer on loandata.customer = customer.ID
                                where Device = " + DevID + " order by Collected desc";

            DGV_DHHistory.DataSource = GetDataSet(conString, SQLString);

            int AgreedRetnDateColID = FindColID(DGV_DHHistory, "ReturnDate");
            int ActualRetnDateColID = FindColID(DGV_DHHistory, "Returned");
            int ActualCollDateColID = FindColID(DGV_DHHistory, "Collected");


            foreach (DataGridViewRow row in DGV_DHHistory.Rows)
                try
                {
                    if (Convert.ToDateTime(row.Cells[ActualRetnDateColID].Value) == Convert.ToDateTime(row.Cells[ActualCollDateColID].Value))
                    {
                        row.DefaultCellStyle.BackColor = Color.Beige;
                    }
                }
                catch { }

        }

        public void ROSLoad_Tree()
        {

            //get permissions code

            string PermissionsText = ROSDBQuery.DBSingleRead(@"select Permissions.PermissionText from Permissions
                                                            left join Users on Users.PermissionGroup = Permissions.PermissionID
                                                            Where Username = '"+ OperaterID + "'");

            if (PermissionsText == "")
            {
                ROSDBQuery.DBWrite("insert into Users (UserName,PermissionGroup) Values ('" + OperaterID + "',1)");

                PermissionsText = ROSDBQuery.DBSingleRead(@"select Permissions.PermissionText from Permissions
                                                            left join Users on Users.PermissionGroup = Permissions.PermissionID
                                                            Where Username = '" + OperaterID + "'");
            }

            //Check For admin

            if(PermissionsText.IndexOf("999") == -1)
            {
                buttonROSAdmin.Visible = false;
            }
            else
            {
                buttonROSAdmin.Visible = true;
            }

            string CurrentSelected = "1";
            try
            {
                CurrentSelected = treeViewROS.SelectedNode.Name.ToString();
            }
            catch
            {

            }
            DataTable ROSData = new DataTable();
            DataTable ROSCatchParents = new DataTable();
            DataTable sortedDT = new DataTable();

            ROSCatchParents.Columns.Add("KeyID");
            ROSCatchParents.Columns.Add("Name");
            ROSCatchParents.Columns.Add("Parent");
            ROSCatchParents.Columns["Parent"].DataType = Type.GetType("System.Int32");


            ROSData = ROSDBQuery.DBGetDataSet("Select * from Keys where (tags LIKE '%" + textBoxROSSearch.Text + "%' or Name LIKE '%" + textBoxROSSearch.Text + "%') AND (Master in ("+ PermissionsText + "))");

            treeViewROS.Nodes.Clear();
            /*
            foreach (DataRow row in ROSData.Rows)
            {

                if (textBoxROSSearch.Text == "")
                {
                    //MessageBox.Show(row["Name"].ToString());
                    if (row["Parent"].ToString() == "0")
                    {
                        treeViewROS.Nodes.Add(row["KeyID"].ToString(), row["Name"].ToString());
                    }
                    else
                    {
                        string a = row["Name"].ToString();
                        string b = row["KeyID"].ToString();
                        string c = row["Parent"].ToString();

                        try
                        {
                            treeViewROS.Nodes[row["Parent"].ToString()].Nodes.Add(row["KeyID"].ToString(), row["Name"].ToString());
                        }
                        catch
                        {
                            try
                            {
                                TreeNode[] tt = this.treeViewROS.Nodes.Find(row["Parent"].ToString(), true);
                                this.treeViewROS.SelectedNode = tt[0];

                                treeViewROS.SelectedNode.Nodes.Add(row["KeyID"].ToString(), row["Name"].ToString());
                            }
                            catch
                            {

                            }
                        }
                    }
                }
                else
                {
                    treeViewROS.Nodes.Add(row["KeyID"].ToString(), row["Name"].ToString());
                }
                
            }
            try
            {
                treeViewROS.CollapseAll();
                TreeNode[] tt2 = this.treeViewROS.Nodes.Find(CurrentSelected, true);
                this.treeViewROS.SelectedNode = tt2[0];
            } catch { }
            //*/


            // Version 2 to tree searched results. Using a second datatable to store failed reseults.
            foreach (DataRow row in ROSData.Rows)
            {
                string ROSName = row["Name"].ToString();
                string ROSKeyID = row["KeyID"].ToString();
                string ROSParent = row["Parent"].ToString();
                //MessageBox.Show(ROSKeyID + ", " + ROSName + " : " + ROSParent);
                try
                {
                    treeViewROS.Nodes[row["Parent"].ToString()].Nodes.Add(row["KeyID"].ToString(), row["Name"].ToString());
                }
                catch
                {
                    
                        while (true)
                        {

                        //MessageBox.Show(ROSKeyID + ", " + ROSName + " : " + ROSParent);
                        try
                        {
                            if (ROSParent == "0")
                            {
                                //MessageBox.Show("Adding Parent: " + ROSKeyID + "/"+ ROSName);
                                treeViewROS.Nodes.Add(ROSKeyID, ROSName);
                            }
                            else
                            {
                                //MessageBox.Show("Adding Node: " + ROSKeyID + "/" + ROSName + "/" + ROSParent);
                                TreeNode[] tt = this.treeViewROS.Nodes.Find(ROSParent, true);
                                //MessageBox.Show("" + tt[0].ToString());
                                this.treeViewROS.SelectedNode = tt[0];

                                treeViewROS.SelectedNode.Nodes.Add(ROSKeyID, ROSName);

                            }

                            sortedDT.Clear();
                            ROSCatchParents.DefaultView.Sort = "Parent asc";
                            sortedDT = ROSCatchParents.DefaultView.ToTable();

                            //MessageBox.Show("New Search");

                            foreach (DataRow ParentRow in sortedDT.Rows)
                            {
                                //MessageBox.Show("From Temp DB - " + ParentRow["Parent"].ToString() + ", " + ParentRow["KeyID"].ToString() + ", " + ParentRow["Name"].ToString());

                                TreeNode[] tt = this.treeViewROS.Nodes.Find(ParentRow["Parent"].ToString(), true);
                                this.treeViewROS.SelectedNode = tt[0];

                                treeViewROS.SelectedNode.Nodes.Add(ParentRow["KeyID"].ToString(), ParentRow["Name"].ToString());
                                //*/

                            }

                            ROSCatchParents.Clear();



                            break;
                        }
                        catch
                        {
                            //MessageBox.Show(ROSKeyID);

                            if (ROSKeyID == "")
                            {
                                break;
                            }
                            else
                            {


                                //MessageBox.Show("Adding to temp db " + ROSKeyID + "/" + ROSName + "/" + ROSParent);
                            ROSCatchParents.NewRow();
                            ROSCatchParents.Rows.Add(ROSKeyID, ROSName, ROSParent);


                            ROSKeyID = ROSParent;

                            ROSParent = ROSDBQuery.DBSingleRead("Select parent from Keys where KeyID = " + ROSKeyID);
                            ROSName = ROSDBQuery.DBSingleRead("Select Name from Keys where KeyID = " + ROSKeyID);
                        }
                  }

                            
          }
                        
      }








            }

            try
            {
                treeViewROS.CollapseAll();
                TreeNode[] tt2 = this.treeViewROS.Nodes.Find(CurrentSelected, true);
                this.treeViewROS.SelectedNode = tt2[0];
            }
            catch { }
        }


        private void buttonROSAddNotes_Click(object sender, EventArgs e)
        {
            AddNewNote();
        }

        public void AddNewNote()
        {
            textBoxROSNotes2.Text = textBoxROSNewNotes.Text + Environment.NewLine + Environment.NewLine + textROSBoxNotes.Text;
            ROSDBQuery.DBWrite("insert into Notes (Key, Note) values ('" + treeViewROS.SelectedNode.Name.ToString() + "','(" + OperaterID + " " + DateTime.Now + ") " + textBoxROSNewNotes.Text + "')");
            textBoxROSNewNotes.Text = "";
        }

        private void buttonROSHideDetails_Click(object sender, EventArgs e)
        {
            panelROSDetails.Visible = false;
            panelROSMain.Visible = true;
            panelROSMain.Dock = DockStyle.Fill;
            ROSLoad_Tree();
        }

        private void ButtonROSRefresh_Click(object sender, EventArgs e)
        {
            ROSLoad_Tree();
        }

        private void buttonROSCopy_Click(object sender, EventArgs e)
        {
            if (buttonROSCopy.Text == "Copy")
            {
                Clipboard.SetText(textBoxROSDisplayKey.Text);
            }
            if (buttonROSCopy.Text == "Add")
            {
                ROSDBQuery.DBWrite("update Keys SET Key = '" + ENC.Encrypt(textBoxROSDisplayKey.Text) + "' where KeyID = " + treeViewROS.SelectedNode.Name.ToString());
                ROSLoad_Tree();
            }
        }

        private void textBoxNotes_DoubleClick(object sender, EventArgs e)
        {
            if (treeViewROS.SelectedNode.Name.ToString() == "0") { return; }

            panelROSMain.Visible = false;
            panelROSDetails.Visible = true;
            panelROSDetails.Dock = DockStyle.Fill;

            textBoxROSTags.Text = ROSDBQuery.DBSingleRead("select Tags from Keys where KeyID = " + treeViewROS.SelectedNode.Name.ToString());

        }

        private void buttonROSNew_Click(object sender, EventArgs e)
        {
            AddNewItem();
                    }

        public void AddNewItem()
        {
            if (textBoxROSNewEntry.Text != "")
            {
                string Parent = "0";
                try
                {
                    Parent = treeViewROS.SelectedNode.Name.ToString();
                }
                catch
                {
                    Parent = "0";
                }
                string Master = ROSDBQuery.DBSingleRead("select Master from Keys where KeyID = " + Parent);
                ROSDBQuery.DBWrite(@"insert into keys (Keystyle, Parent, Name, Tags, Master) values (2," + Parent + ",'" + textBoxROSNewEntry.Text + "','',"+ Master + ")");
                textBoxROSNewEntry.Clear();
                ROSLoad_Tree();
            }
        }

        private void treeViewROS_AfterSelect(object sender, TreeViewEventArgs e)
        {
            labelROSID.Text = treeViewROS.SelectedNode.Name.ToString();

            try
            {
                textBoxROSDisplayKey.Text = ENC.Decrypt(ROSDBQuery.DBSingleRead("select key from Keys where KeyID = " + treeViewROS.SelectedNode.Name.ToString() + ""));
                buttonROSCopy.Text = "Copy";
                textBoxROSDisplayKey.Enabled = false;
            }
            catch
            {
                textBoxROSDisplayKey.Text = "";
                buttonROSCopy.Text = "Add";
                textBoxROSDisplayKey.Enabled = true;
            }

            labelROSSelected.Text = ROSDBQuery.DBSingleRead("select Name from Keys where KeyID = " + treeViewROS.SelectedNode.Name.ToString());
            DataTable Notes = ROSDBQuery.DBGetDataSet("select Note from Notes where Key = " + treeViewROS.SelectedNode.Name.ToString() + " order by NoteID Desc");
            
            buttonROSNew.Text = "Add New to " + labelROSSelected.Text;

            if (textBoxROSDisplayKey.Text == "" || textBoxROSDisplayKey.Text == "0")
            {
                buttonROSNew.Enabled = true;
            }
            else
            {
                buttonROSNew.Enabled = false;
            }

            string CompressedString = "";

            foreach (DataRow dr in Notes.Rows)
            {
                CompressedString = CompressedString + dr["Note"].ToString() + Environment.NewLine + Environment.NewLine;
            }

            textROSBoxNotes.Text = CompressedString;
            textBoxROSNotes2.Text = CompressedString;
        }

        private void buttonROSAddTag_Click(object sender, EventArgs e)
        {
            AddNewTag();
        }

        public void AddNewTag()
        {
            ROSDBQuery.DBWrite("update Keys set Tags = '" + textBoxROSTags.Text + ' ' + textBoxROSNewTag.Text + "' where KeyID = " + treeViewROS.SelectedNode.Name.ToString());
            textBoxROSTags.Text = ROSDBQuery.DBSingleRead("select Tags from Keys where KeyID = " + treeViewROS.SelectedNode.Name.ToString());
            textBoxROSNewTag.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBoxROSSearch.Text = "";
            ROSLoad_Tree();
        }

        private void textBoxROSSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                e.Handled = true;
                ROSLoad_Tree();
            }
        }

        private void textBoxROSNewEntry_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                e.Handled = true;
                AddNewItem();
            }
        }

        private void textBoxROSNewNotes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                e.Handled = true;
                AddNewNote();
            }
        }

        private void textBoxROSNewTag_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            if (e.KeyChar == (char)Keys.Return)
            {
                e.Handled = true;
                AddNewTag();
            }
        }

        private void buttonROSAdmin_Click(object sender, EventArgs e)
        {
            panelROSMain.Visible = false;
            panelROSAdmin.Visible = true;
            panelROSAdmin.Dock = DockStyle.Fill;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            RefreshAdmin();
        }

        public void RefreshAdmin()
        {
            DGVROSGroups.AutoGenerateColumns = true;
            DGVROSGroups.DataSource = ROSDBQuery.DBGetDataSet("select * from Permissions");

            DGVROSGroups.Columns["PermissionID"].Visible = false;
            DGVROSGroups.Columns["PermissionCode"].Visible = false;
            DGVROSGroups.Columns["PermissionText"].Visible = false;

            DGVROSPermissions.Columns.Clear();
            DGVROSPermissions.DataSource = ROSDBQuery.DBGetDataSet("select KeyID, Parent, Name, Master from Keys where KeyID < 26");

            DGVROSPermissions.Columns["KeyID"].Visible = false;
            DGVROSPermissions.Columns["Parent"].Visible = false;
            DGVROSPermissions.Columns["Master"].Visible = false;

            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "Permission";
            checkColumn.HeaderText = "Permission";
            checkColumn.Width = 125;
            checkColumn.ReadOnly = false;
            checkColumn.FillWeight = 125; //if the datagridview is resized (on form resize) the checkbox won't take up too much; value is relative to the other columns' fill values
            DGVROSPermissions.Columns.Add(checkColumn);

            //Users Box
            DGVROSUsers.Columns.Clear();
            DGVROSUsers.DataSource = ROSDBQuery.DBGetDataSet("select * from Users");

            DataGridViewComboBoxColumn comboboxinput = new DataGridViewComboBoxColumn();
            comboboxinput.Name = "Assigned Group";
            comboboxinput.HeaderText = "Assigned Group";
            comboboxinput.Width = 125;
            comboboxinput.ReadOnly = false;
            comboboxinput.FillWeight = 125;

            comboboxinput.DataSource = ROSDBQuery.DBGetDataSet("select * from Permissions");
            comboboxinput.DisplayMember = "GroupName";
            comboboxinput.ValueMember = "PermissionID";

            DGVROSUsers.Columns.Add(comboboxinput);

            DGVROSUsers.Columns["UserID"].Visible = false;
            DGVROSUsers.Columns["PermissionGroup"].Visible = false;

            foreach (DataGridViewRow row in DGVROSUsers.Rows)
            {
                row.Cells["Assigned Group"].Value = row.Cells["PermissionGroup"].Value;
            }
            buttonROSGroupSave.Visible = false;
            buttonROSPermissionSave.Visible = false;
        }


    private void buttonRosHideAdmin_Click(object sender, EventArgs e)
        {
            panelROSAdmin.Visible = false;
            panelROSMain.Visible = true;
            panelROSMain.Dock = DockStyle.Fill;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ROSDBQuery.DBWrite("Insert into Permissions (GroupName, PermissionCode,PermissionText) Values ('" + textBoxRosNewGroup.Text + "','010000000000000000000000110','101,125')"); //First isnt read but is needed to offset
            textBoxRosNewGroup.Text = "";
            RefreshAdmin();

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (textBoxRosNewMaster.Text == "")
            {
                return;
            }

            int NewID = Convert.ToInt32(ROSDBQuery.DBSingleRead("select count (*) from Keys where KeyID < 24"));

            if (NewID == 23)
            {
                return;
            }

            NewID = NewID + 1;
            int NewMaster = NewID + 100;


            ROSDBQuery.DBWrite("Insert into Keys (KeyID,KeyStyle,Parent,Name,Master) Values (" + NewID + ",1,0,'" + textBoxRosNewMaster.Text + "'," + NewMaster + ")");
            textBoxRosNewMaster.Text = "";
            RefreshAdmin();
        }


        

        private void DGVROSGroups_Click(object sender, EventArgs e)
        {
            string PermissionsCode = DGVROSGroups.CurrentRow.Cells["PermissionCode"].Value.ToString();


            
            foreach (DataGridViewRow row in DGVROSPermissions.Rows)
            {
                int KeyID = Convert.ToInt32( row.Cells["KeyID"].Value);

                //MessageBox.Show(KeyID.ToString());

                row.Cells["Permission"].Value = Convert.ToInt32(PermissionsCode.Substring(KeyID, 1));

               
            }

            buttonROSPermissionSave.Visible = false;

        }

        private void DGVROSUsers_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            buttonROSGroupSave.Visible = true;
        }

        private void DGVROSUsers_CellStateChanged(object sender, DataGridViewCellStateChangedEventArgs e)
        {
            buttonROSGroupSave.Visible = true;
        }

        private void buttonROSGroupSave_Click(object sender, EventArgs e)
        {
            buttonROSGroupSave.Visible = false;
            foreach (DataGridViewRow row in DGVROSUsers.Rows)
            {
                //row.Cells["Assigned Group"].Value = row.Cells["PermissionGroup"].Value;
                //MessageBox.Show(row.Cells["UserID"].Value + " - " + row.Cells["Assigned Group"].Value);
                ROSDBQuery.DBWrite("update Users SET PermissionGroup = '" + row.Cells["Assigned Group"].Value + "' where UserID = " + row.Cells["UserID"].Value);
            }


        }

        private void DGVROSPermissions_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            buttonROSPermissionSave.Visible = true;
        }
        private void DGVROSPermissions_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            buttonROSPermissionSave.Visible = true;
        }

        private void buttonROSPermissionSave_Click(object sender, EventArgs e)
        {

            StringBuilder sb = new StringBuilder(DGVROSGroups.CurrentRow.Cells["PermissionCode"].Value.ToString());

            //MessageBox.Show(sb.ToString());

            foreach (DataGridViewRow row in DGVROSPermissions.Rows)
            {

                //MessageBox.Show(Convert.ToInt32(row.Cells["KeyID"].Value).ToString() + " - " + row.Cells["Permission"].Value.ToString());

                int KeyID = Convert.ToInt32(row.Cells["KeyID"].Value);

                if (row.Cells["Permission"].Value.ToString() == "True" || row.Cells["Permission"].Value.ToString() == "1")
                {
                    sb[KeyID] = '1'; // index starts at 0!
                }
                else
                {
                    sb[KeyID] = '0'; // index starts at 0!

                }


            }

            //MessageBox.Show(sb.ToString());
            //MessageBox.Show(DGVROSGroups.CurrentRow.Cells["PermissionID"].Value.ToString());
            ROSDBQuery.DBWrite("update Permissions SET PermissionCode = '" + sb.ToString() + "' where PermissionID = " + DGVROSGroups.CurrentRow.Cells["PermissionID"].Value.ToString());

            string PermissionsText = "";
            int i = 0;
            foreach (char c in sb.ToString())
            {

                if (c.ToString() == "1")
                {
                    //MessageBox.Show(i.ToString());
                    if (i == 26)
                    {
                        PermissionsText = PermissionsText + "999,";
                    }
                    else
                    {
                        PermissionsText = PermissionsText + (i + 100) + ",";
                    }
                }



                i++;
            }
            PermissionsText = PermissionsText + "0";  //added so last value is not a comma
            //MessageBox.Show(PermissionsText);
            ROSDBQuery.DBWrite("update Permissions SET PermissionText = '" + PermissionsText + "' where PermissionID = " + DGVROSGroups.CurrentRow.Cells["PermissionID"].Value.ToString());
            buttonROSPermissionSave.Visible = false;
        }
    }
}



    
        
    
