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
using System.DirectoryServices;
using System.IO;
using System.Net;
using System.DirectoryServices.ActiveDirectory;
using Microsoft.VisualBasic;
using Outlook = Microsoft.Office.Interop.Outlook;










namespace UoN_App_SQLLite
{

    

    public partial class FormLoans : Form
    {
        Int32 SlideMove = 20;
        Int32 SlideOpen = 450;
        Int32 SlideClose = 40;
        Int32 SlideInterval = 20;
        Boolean ReadCustomersbackgroundWorkerRunAgain = false;
        Color currentColor = Color.Green;


        DateTime CollectionTime = Convert.ToDateTime("08:15:00");
        DateTime ReturnTime = Convert.ToDateTime("08:00:00");

        public void EnableDoubleBuffering()
        {
            // Set the value of the double-buffering style bits to true.
            this.SetStyle(ControlStyles.DoubleBuffer |
               ControlStyles.UserPaint |
               ControlStyles.AllPaintingInWmPaint,
               true);
            this.UpdateStyles();
        }

        public static int GetMonthDifference(DateTime startDate, DateTime endDate)
        {
            int monthsApart = 365 * (startDate.Year - endDate.Year) + 365 * (startDate.Month - endDate.Month);
            return Math.Abs(monthsApart);
        }

        public FormLoans()
        {


            EnableDoubleBuffering();
            InitializeComponent();
            ReadTodaysBookings();

            G1timer.Interval = SlideInterval;
            G2timer.Interval = SlideInterval;
            G3timer.Interval = SlideInterval;


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

        private void LoanForm_Load(object sender, EventArgs e)
        {
            
            Screen screen = Screen.PrimaryScreen;
            //int S_width = screen.Bounds.Width;
            //int S_height = screen.Bounds.Height;
            //this.MaximumSize = new System.Drawing.Size(700, 600);

            this.Width = 700;
            this.Height = 600;

            G1groupBox.Top = 10;
            G2groupBox.Top = 470;
            G3groupBox.Top = 510;
            G1groupBox.Anchor = (AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top);

            G1groupBox.Top = 10;
            G1groupBox.Left = 22;

            //G2groupBox.Top = 510;
            G2groupBox.Top = 210;
            G2groupBox.Left = 22;
            G2groupBox.Height = 30;
            G2groupBox.Anchor = (AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top);

            //G3groupBox.Top = 590;
            G3groupBox.Top = 310;
            G3groupBox.Left = 22;
            G3groupBox.Height = 30;
            G3groupBox.Anchor = (AnchorStyles.Right | AnchorStyles.Left | AnchorStyles.Top);



            WaitpictureBox.Top = 300;
            WaitpictureBox.Left = 140;

            EmailConfirmlabel.Top = 350;
            EmailConfirmlabel.Left = 160;
            //*/



            ReadData();
            G1timer.Enabled = true;
            this.AutoScaleMode = AutoScaleMode.Dpi;

            /*
            Timer timer = new Timer();
            timer.Interval = (10 * 1000); // 10 secs
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();
            */

        }

        private void timer_Tick(object sender, EventArgs e)
        {


        }

        private void FormLoans_Deactivate(object sender, EventArgs e)
        {
            
            
        }

        private void Refreshtimer_Tick(object sender, EventArgs e)
        {
            

            Form fc = Application.OpenForms["FormLoans"];

            if (fc != null)
            {
                ReadTodaysBookings();
                
            }
            else
            {
                //MessageBox.Show("Terminate Refresh");
                Refreshtimer.Enabled = false;
            }
                
            
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

        void ReadData()
        {
            var DBSearchString = "";

            if (textBox1.Text != "")
            {
                var SearchString = textBox1.Text.Replace("'", "''");
                DBSearchString = " WHERE s LIKE '%" + SearchString + "%'";

            }

            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    Conn.Open();

                    /*

                    //ORDER BY column1, column2, ... ASC|DESC
                    cmd.CommandText = "SELECT * FROM LoanData ORDER BY Id ASC";

                    LoanDataView1.Rows.Clear();

                    using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            LoanDataView1.Rows.Add(new object[]
                            {
                            reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("ID")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("Device")),
                                    reader.GetValue(reader.GetOrdinal("Customer")),
                                    reader.GetValue(reader.GetOrdinal("Assigned")),
                                    reader.GetValue(reader.GetOrdinal("CreatedBy")),
                                    reader.GetValue(reader.GetOrdinal("CollectionDate")),
                                    reader.GetValue(reader.GetOrdinal("ReturnDate")),
                                    reader.GetValue(reader.GetOrdinal("Notes")),
                                    reader.GetValue(reader.GetOrdinal("Delivery")),
                                    reader.GetValue(reader.GetOrdinal("Returned"))


                            });
                        }
                    }
                    cmd.CommandText = "SELECT * FROM Admins ORDER BY Surname ASC";

                    AdminsView1.Rows.Clear();

                    using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            AdminsView1.Rows.Add(new object[]
                            {
                            reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("ID")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("FirstName")),
                                    reader.GetValue(reader.GetOrdinal("Surname")),
                                    reader.GetValue(reader.GetOrdinal("Permissions")),
                                    reader.GetValue(reader.GetOrdinal("Telephone")),
                                    reader.GetValue(reader.GetOrdinal("Email")),
                                    reader.GetValue(reader.GetOrdinal("Active"))

                            });
                        }
                    }

                    */

                    cmd.CommandText = "SELECT * FROM LoanDescriptions ORDER BY SortOrder ASC";

                    LoanDescriptionView1.Rows.Clear();

                    

                    using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            LoanDescriptionView1.Rows.Add(new object[]
                            {
                            reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("ID")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("Description")),
                                    reader.GetValue(reader.GetOrdinal("Description"))

                            });
                        }
                    }

                    cmd.CommandText = "SELECT * FROM LoanLocations ORDER BY ID ASC";

                    LoanLocationsView1.Rows.Clear();

                    this.LoanLocationsView1.Rows.Add("0", "0", "Please Select a Location");

                    using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            LoanLocationsView1.Rows.Add(new object[]
                            {
                            reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("ID")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("Location"))

                            });
                        }
                    }

                    Conn.Close();

                    ReturndateTimePicker.Value = ReturndateTimePicker.Value;

                    CustomersRead();
                    //AvailableStockRead();
                }
            }
        }

        private void SendEmail(String UserEmail)
        {
            // *******  DECREPID ******** Code replaced 18/08/17 to include UoN Logo.
            //MessageBox.Show("Token number is: " + System.Security.Principal.WindowsIdentity.GetCurrent().Token);


            MailMessage msg = new MailMessage();
            msg.To.Add(new MailAddress(UserEmail));
            msg.From = new MailAddress("AVBookings@northampton.ac.uk");
            msg.Subject = "Your Equipment Booking Confirmation";
            //msg.Body = "<div>This is a HTML email test.</div>";

            string FullEmail = File.ReadAllText(Directory.GetCurrentDirectory() + $"/Email-Start.html");


            try
            {

                // Generate middle email from CusromerOrders Gridview.

                foreach (DataGridViewRow row in CustomerOrder.Rows)
                    if (row.Cells[1].Value.ToString() != "")
                    {
                        //MessageBox.Show("" + row.Cells[0].Value.ToString());
                        FullEmail = FullEmail + "<tr><td>" + row.Cells[0].Value.ToString() + "</td><td>" + row.Cells[1].Value.ToString() + "</td><td>" + row.Cells[2].Value.ToString() + "</td><td>" + row.Cells[3].Value.ToString() + "</td><td>" + row.Cells[4].Value.ToString() + "</td></tr>";
                    }
            }
            catch { }

            //Assemble full email.

            FullEmail = FullEmail + File.ReadAllText(Directory.GetCurrentDirectory() + $"/Email-End.html");

            msg.Body = FullEmail;

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

            htmlView.LinkedResources.Add(LinkedImage);
            msg.AlternateViews.Add(htmlView);

            msg.IsBodyHtml = true;


            SmtpClient client = new SmtpClient();
            client.Host = "webmail.northampton.ac.uk";

            client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

            client.Port = 25;
            client.EnableSsl = false;
            client.UseDefaultCredentials = true;

            //MessageBox.Show("" + System.Net.CredentialCache.DefaultNetworkCredentials.UserName);

            client.Send(msg);

            //*/
        }

        private void SendEmailWithLogo(String UserEmail, String FullName)
        {
            //MessageBox.Show("" + UserID);
            //UserEmail = "Simon.Ford@northampton.ac.uk";
            //UserEmail = "mark.rowland@northampton.ac.uk";
            //UserEmail = "James.Gough@northampton.ac.uk";
            //UserEmail = "INS_Service_Desk_Team@northampton.ac.uk";

            //MessageBox.Show("Token number is: " + System.Security.Principal.WindowsIdentity.GetCurrent().Token);

            Outlook.Application oApp = new Outlook.Application();
            Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

            //MailMessage msg = new MailMessage();
            oMailItem.To = UserEmail;
            //msg.IsBodyHtml = true;
            //msg.From = new MailAddress("AVBookings@northampton.ac.uk");
            oMailItem.Subject = "Equipment booking confirmation";

            
            string FullEmail = @"<h1 style=""text - align: center; font - family: OpenSans; color: #5e9ca0;"">Equipment Booking Confirmation with IT Services.</h1>
<p></p>
      <h2 style=""color: #2e6c80; font-family: OpenSans;"" > Dear " + FullName + ",<p></p>Thank you for placing your order with the University of Northampton's IT Services.</h2>";


            FullEmail = FullEmail + File.ReadAllText(Directory.GetCurrentDirectory() + $"/Email-Start.html");




                // Generate middle email from CusromerOrders Gridview.

                try
                {


                    
                    // Generate middle email from CusromerOrders Gridview.

                    foreach (DataGridViewRow row in CustomerOrder.Rows)
                        if (row.Cells[1].Value.ToString() != "")
                        {
                            //MessageBox.Show("" + row.Cells[0].Value.ToString());
                            FullEmail = FullEmail + "<tr><td>" + row.Cells[0].Value.ToString() + "</td><td>" + row.Cells[1].Value.ToString() + "</td><td>" + row.Cells[2].Value.ToString() + "</td><td>" + row.Cells[3].Value.ToString() + "</td><td>" + row.Cells[4].Value.ToString() + "</td></tr>";
                        }
                }
                catch { }

                //Assemble full email.

                FullEmail = FullEmail + File.ReadAllText(Directory.GetCurrentDirectory() + $"/Email-End.html");

                // Add logo and contact info.

                LinkedResource LinkedImage = new LinkedResource(Environment.CurrentDirectory + @"\EmailLogo.png");
                LinkedImage.ContentId = "MyPic";

                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(FullEmail +
      @"<img src=cid:MyPic>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong><a href=""http://www.northampton.ac.uk/unit/"">Northampton.ac.uk</a></p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong  >University of Northampton,</strong> Waterside Campus,</p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;"">Waterside Campus, University Drive, Northampton, NN1 5PH</p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><strong>Follow the story on social media</strong></p>
            <p style=""margin: 0in; font - family: OpenSans; font - size: 12.0pt; color: #333333;""><a href=""http://www.northampton.ac.uk/social-media-hub/"">http://www.northampton.ac.uk/social-media-hub/</a></p>
                ",
      null, "text/html");

                htmlView.LinkedResources.Add(LinkedImage);

            /*
                SmtpClient client = new SmtpClient();
                client.Host = "webmail.northampton.ac.uk";

                client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

                client.Port = 25;
                client.EnableSsl = false;
                client.UseDefaultCredentials = true;
            */
            //MessageBox.Show("" + System.Net.CredentialCache.DefaultNetworkCredentials.UserName);

            //MessageBox.Show("Send Email");

            //client.Send(msg);
            oMailItem.HTMLBody = FullEmail;
            oMailItem.Display(false);


            //*/

        }

        private void SendCalendarInvite()
        {


            //MessageBox.Show("Wait");

            foreach (DataGridViewRow row in CustomerOrder.Rows)
                try
                {
                    if (row.Cells[1].Value.ToString() != "")
                    {
                        DateTime startTime = Convert.ToDateTime(row.Cells[2].Value.ToString());
                        DateTime endTime = Convert.ToDateTime(row.Cells[4].Value.ToString());
                        String AppointmentTitle = row.Cells[1].Value.ToString() + " " + label3.Text;
                        String Campus = row.Cells[3].Value.ToString();

                        //MessageBox.Show("" + row.Cells[0].Value.ToString());
                        //FullEmail = FullEmail + "<tr><td>" + row.Cells[0].Value.ToString() + "</td><td>" + row.Cells[1].Value.ToString() + "</td><td>" + row.Cells[2].Value.ToString() + "</td><td>" + row.Cells[3].Value.ToString() + "</td><td>" + row.Cells[4].Value.ToString() + "</td></tr>";


                        //DateTime startTime = Convert.ToDateTime("2017/06/05 11:30:00");
                        //DateTime endTime = Convert.ToDateTime("2017/06/05 12:30:00");

                        SmtpClient sc = new SmtpClient();
                        MailMessage msg = new MailMessage();
                        msg.From = new MailAddress("AVBookings@northampton.ac.uk");
                        msg.To.Add(new MailAddress("LoanCalendar@northampton.ac.uk"));

                        msg.Subject = AppointmentTitle;
                        msg.Body = "Automatic email from the UoN Loan App.";

                        StringBuilder str = new StringBuilder();
                        str.AppendLine("BEGIN:VCALENDAR");
                        //str.AppendLine("PRODID:-//Ahmed Abu Dagga Blog");
                        str.AppendLine("VERSION:2.0");
                        str.AppendLine("METHOD:REQUEST");
                        str.AppendLine("BEGIN:VEVENT");
                        str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", startTime));
                        str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
                        str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", endTime));
                        str.AppendLine("LOCATION: " + Campus);
                        str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
                        str.AppendLine(string.Format("DESCRIPTION:{0}", msg.Body));
                        str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", msg.Body));
                        str.AppendLine(string.Format("SUMMARY:{0}", msg.Subject));
                        str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));

                        str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", msg.To[0].DisplayName, msg.To[0].Address));

                        str.AppendLine("BEGIN:VALARM");
                        str.AppendLine("TRIGGER:-PT15M");
                        str.AppendLine("ACTION:DISPLAY");
                        str.AppendLine("DESCRIPTION:Reminder");
                        str.AppendLine("END:VALARM");
                        str.AppendLine("END:VEVENT");
                        str.AppendLine("END:VCALENDAR");
                        System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("text/calendar");
                        ct.Parameters.Add("method", "REQUEST");
                        AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), ct);
                        msg.AlternateViews.Add(avCal);

                        sc.Host = "webmail.northampton.ac.uk";
                        sc.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

                        //MessageBox.Show("Check?");

                        sc.Send(msg);
                    }
                }
                catch { }
                

        }

        void AvailableStockRead()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.AutoGenerateColumns = true;

            string DTNow = DateTime.Now.ToString("yyyy-MM-dd");
            string CollectionDate = CollectdateTimePicker.Value.ToString("yyyy-MM-dd");
            string ReturnDate = ReturndateTimePicker.Value.ToString("yyyy-MM-dd");
            string LoanStockDesc = LoanDescriptionView1.CurrentRow.Cells[1].Value.ToString();
            string LocDesc = LoanLocationsView1.CurrentRow.Cells[1].Value.ToString();

            //exit if not required.
            if (CollectionDate == "" || ReturnDate == "" || LoanStockDesc == "" || LocDesc == "0")
            {    
                return;
            }

            

            //try
            //{

            

                //MessageBox.Show(LoanLocationsView1.CurrentRow.Cells[1].Value.ToString());
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                        //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)

                        if (CollectionDate != DTNow)
                        {
                            cmd.CommandText = @"
                                SELECT DISTINCT LoanStock.AssetID, LoanDescriptions.Description, LoanLocations.Location ,LoanStock.ID 
                                FROM LoanStock INNER JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID 
                                INNER JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID CROSS JOIN LoanData 
                                WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + CollectionDate + " 10:00:00') AND(CollectionDate <= '" + ReturnDate + @" 18:00:00') AND LoanData.Returned IS NULL))) 
                                AND (LoanStock.Description = " + LoanStockDesc + ")  AND (LoanStock.Location = " + LocDesc + ") AND LoanStock.Active = 1 ORDER BY LoanStock.AssetID ASC";
                            //cmd.CommandText = "SELECT DISTINCT LoanStock.AssetID, LoanDescriptions.Description, LoanLocations.Location ,LoanStock.ID FROM LoanStock INNER JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID INNER JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID CROSS JOIN LoanData WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + CollectdateTimePicker.Value.ToString("yyyy-MM-dd") + " 13:00:00') AND (CollectionDate <= '2017-05-30 11:00:00') AND LoanData.Returned IS NULL AND LoanData.Collected IS NULL))) AND (LoanStock.Description = 1)  AND (LoanStock.Location = 1) AND LoanStock.Active = 1 ORDER BY LoanStock.ID ASC";
                        }
                        else
                        {
                            cmd.CommandText = @"
                                SELECT DISTINCT LoanStock.AssetID, LoanDescriptions.Description, LoanLocations.Location ,LoanStock.ID
                                FROM LoanStock INNER JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID
								INNER JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID CROSS JOIN LoanData
								WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + ReturnDate + " 10:00:00') AND (CollectionDate <= '" + CollectionDate + @" 18:00:00') AND LoanData.Returned IS NULL))) 
                                AND  
								NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE Returndate < '" + ReturnDate + @" 08:00:00' And Returned is null))
                                AND (LoanStock.Description = " + LoanStockDesc + ")  AND (LoanStock.Location = " + LocDesc + ") AND LoanStock.Active = 1 ORDER BY LoanStock.AssetID ASC";
                    }

                        
                        dataGridView1.DataSource = DBGetDataSet(cmd.CommandText);

                        /*
                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                dataGridView1.Rows.Add(new object[]
                                {
                                    //reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("ID")),
                                    reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("Description")),
                                    reader.GetValue(reader.GetOrdinal("Location"))



                                });
                            }
                        }
                        */

                    }

                    Conn.Close();


                }
            /*}
            catch (Exception err)
            {
                HandleError("AvailableStockRead() : " + err);
            }
            */
        }

        void ReadTodaysBookings()
        {
            /*
            try
            {
            //*/
                //MessageBox.Show(LoanLocationsView1.CurrentRow.Cells[1].Value.ToString());
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                        //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)
                        //MessageBox.Show(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (NOT (LoanData.ID IN (SELECT Id FROM LoanData WHERE returned = collected AND returned not null))) AND (CollectionDate > '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 08:00:00' and CollectionDate < '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 17:30:00' AND LoanData.Collected is null) or (CollectionDate < '" + Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 08:00')"); //AND LoanData.Collected = 0
                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (CollectionDate > '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 08:00:00' and CollectionDate < '" + dateTimePicker1.Value.ToString("yyyy-MM-dd" ) + " 17:30:00' AND LoanData.Collected is null) or (CollectionDate < '" + Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 08:00')"); //AND LoanData.Collected = 0
                        cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.ReturnDate, LoanData.Id, Customer.Firstname FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (Collectiondate > '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 08:00:00' And Collectiondate < '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 18:00:00' And Collected is null and returned is null) or (Collectiondate < '" + DateTime.Now.ToString("yyyy-MM-dd") + " 08:00:00' And Collected is null and returned is null) ORDER BY Collectiondate ASC";

                        dataGridView2.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                dataGridView2.Rows.Add(new object[]
                                {
                            reader.GetValue(7) + " " + reader.GetValue(0),  //reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                            reader.GetValue(1),        //reader.GetValue(reader.GetOrdinal("Description")),
                            reader.GetValue(2),        //reader.GetValue(reader.GetOrdinal("Location")),
                            reader.GetValue(3),
                            reader.GetValue(4),
                            reader.GetValue(5),
                            reader.GetValue(6)//reader.GetValue(reader.GetOrdinal("CreatedBy")),
                                    //reader.GetValue(reader.GetOrdinal("CollectionDate"))


                                });
                            }
                        }

                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Returned, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE (NOT (LoanData.ID IN (SELECT Id FROM LoanData WHERE returned = collected AND returned not null))) ReturnDate > '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 01:00:00' AND ReturnDate < '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 23:00:00'  AND LoanData.Collected is null"; //AND LoanData.Returned = 0
                        //cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.Collected, LoanData.Id FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE(Returndate > '2017-06-09 08:00:00' And Returndate < '2017-06-09 17:00:00' And Returned is null) or(Returndate < '"+ DateTime.Now + "' And Returned is null)";
                        cmd.CommandText = "SELECT Customer.Surname, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.ReturnDate, LoanData.Returned, LoanData.Id, Customer.Firstname, Customer.Telephone FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID LEFT JOIN Customer ON LoanData.Customer = Customer.ID WHERE Customer.ID != 1 and (Returndate > '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 08:00:00' And Returndate < '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 18:00:00' And Returned is null) or (Returndate < '" + DateTime.Now.ToString("yyyy-MM-dd") + " 08:00:00' And Returned is null) ORDER BY Returndate ASC";

                        dataGridView3.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                dataGridView3.Rows.Add(new object[]
                                {
                            reader.GetValue(7) + " " + reader.GetValue(0),  // U can use column index
                            reader.GetValue(8),
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

                    //MessageBox.Show("Here");
                    DangerLoandataGridView.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                
                                DangerLoandataGridView.Rows.Add(new object[]
                                {
                                    
                            reader.GetValue(0),  // U can use column index
                            reader.GetValue(1)+ ", " + reader.GetValue(2),        //reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                            reader.GetValue(3),        //reader.GetValue(reader.GetOrdinal("Description")),
                            reader.GetValue(4),        //reader.GetValue(reader.GetOrdinal("Location")),
                            reader.GetValue(5),
                            reader.GetValue(6),
                            reader.GetValue(7),
                            reader.GetValue(8) + ", " + reader.GetValue(9),//reader.GetValue(reader.GetOrdinal("CreatedBy")),
                            reader.GetValue(10)        //reader.GetValue(reader.GetOrdinal("CollectionDate"))


                                });
                            }
                        }



                    }

                    Conn.Close();

                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {

                        //MessageBox.Show("" + Convert.ToDateTime(row.Cells[4].Value.ToString()));

                        if (Convert.ToDateTime(row.Cells[4].Value.ToString()) < Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 01:00"))
                        //dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 23:00:00
                        {
                            row.DefaultCellStyle.BackColor = Color.Orange;
                        }
                    }
                

                    foreach (DataGridViewRow row in dataGridView3.Rows)
                        if (Convert.ToDateTime(row.Cells[5].Value.ToString()) < Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 01:00"))
                        {
                            row.DefaultCellStyle.BackColor = Color.Orange;
                        }

                if (DangerLoandataGridView.RowCount == 0)
                {
                    tabControl2.TabPages[2].Text = "";
                    TabFlashtimer.Stop();

                }
                else
                {
                    //tabControl2.TabPages[2].Text = "ATTENTION!!";
                    TabFlashtimer.Start();
                }

            }
                /*
            }
            catch (Exception err)
            {
                HandleError("ReadTodaysBookings() : " + err);
                //MessageBox.Show("Error..." + err);
            }
    //*/
}

        void CustomerBookingsRead()
        {

            
            try
            {
                //*/
                //MessageBox.Show(LoanLocationsView1.CurrentRow.Cells[1].Value.ToString());
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                    //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)
                        int CustomerIDCol = FindColID(CustomerView1, "ID");
                    
                    
                        
                        cmd.CommandText = "SELECT LoanData.Id, LoanDescriptions.Description, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.ReturnDate, LoanData.Collected, LoanData.Returned, LoanData.ReturnedBy FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID WHERE Customer = " + CustomerView1.CurrentRow.Cells[CustomerIDCol].Value.ToString() + " ORDER BY LoanData.CollectionDate Desc";
                        //MessageBox.Show("hi");
                    

                        UserLoansView.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                UserLoansView.Rows.Add(new object[]
                                {
                            reader.GetValue(0),  // U can use column index
                            reader.GetValue(1),        //reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                            reader.GetValue(2),        //reader.GetValue(reader.GetOrdinal("Description")),
                            reader.GetValue(3),        //reader.GetValue(reader.GetOrdinal("Location")),
                            reader.GetValue(4),
                            reader.GetValue(5),
                            reader.GetValue(6),
                            reader.GetValue(7),
                            reader.GetValue(8)

                                });
                            string a = reader.GetValue(6).ToString();
                            }
                        }
                        
                        cmd.CommandText = "SELECT LoanData.Id, LoanDescriptions.Description as AssetDescription, LoanStock.AssetID, LoanLocations.Location, LoanData.CollectionDate, LoanData.ReturnDate, LoanData.Collected, LoanData.Returned, LoanData.ReleasedBy, LoanStock.Description, LoanStock.ID, LoanData.CollectionPoint FROM LoanData LEFT JOIN LoanStock ON LoanData.Device = LoanStock.ID LEFT JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID WHERE LoanData.Returned IS NULL AND customer = " + CustomerView1.CurrentRow.Cells[CustomerIDCol].Value.ToString() + " ORDER BY LoanData.CollectionDate ASC";
                        //MessageBox.Show("Here");
                        OutStandlingLoans.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                OutStandlingLoans.Rows.Add(new object[]
                                {
                            reader.GetValue(reader.GetOrdinal("Id")),  
                            reader.GetValue(reader.GetOrdinal("AssetDescription")),        //reader.GetValue(reader.GetOrdinal("AssetID")),  // Or column name like this
                            reader.GetValue(reader.GetOrdinal("AssetID")),        //reader.GetValue(reader.GetOrdinal("Description")),
                            reader.GetValue(3),        //reader.GetValue(reader.GetOrdinal("Location")),
                            reader.GetValue(4), //CDate
                            reader.GetValue(5), //RDate
                            reader.GetValue(6), //Collected
                            reader.GetValue(8),
                            reader.GetValue(9),  //Returned
                            reader.GetValue(10),
                            reader.GetValue(11),
                            reader.GetValue(reader.GetOrdinal("CollectionPoint")),
                            reader.GetValue(reader.GetOrdinal("ReleasedBy"))

                                    //reader.GetValue(reader.GetOrdinal("LoanStock.Description"))                    //reader.GetValue(reader.GetOrdinal("CollectionDate"))


                                });
                            }
                        }

                        foreach (DataGridViewRow row in OutStandlingLoans.Rows)
                            if (Convert.ToDateTime(row.Cells[5].Value) < DateTime.Now)
                            {
                                row.DefaultCellStyle.BackColor = Color.Red;
                            }

                        foreach (DataGridViewRow row in UserLoansView.Rows)
                        {
                            
                            {
                                if (Convert.ToString(row.Cells[6].Value) != "" && Convert.ToDateTime(row.Cells[6].Value) > Convert.ToDateTime(row.Cells[5].Value))
                                {
                                    row.DefaultCellStyle.BackColor = Color.Red;
                                }

                                if (Convert.ToString(row.Cells[6].Value) != "" && Convert.ToDateTime(row.Cells[6].Value) < Convert.ToDateTime(row.Cells[5].Value))
                                {
                                    row.DefaultCellStyle.BackColor = Color.LightGreen;
                                }

                                if (Convert.ToString(row.Cells[6].Value) != "" && Convert.ToDateTime(row.Cells[7].Value) == Convert.ToDateTime(row.Cells[6].Value))
                                {
                                    row.DefaultCellStyle.BackColor = Color.Beige;
                                }
                            }
                            
                        }


                    }

                    Conn.Close();


                }
            
            }
            catch
            {
                //HandleError("CustomerBookingsRead() : " + err);
            }
            //*/
        }

        void CustomersRead()
        {
            //MessageBox.Show("");
            try
            {
                
                //MessageBox.Show(LoanLocationsView1.CurrentRow.Cells[1].Value.ToString());
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();

                        

                        string SearchText = textBox1.Text.Trim();

                        bool contains = SearchText.Contains(",");
                        foreach (var sentence in SearchText.TrimEnd('.').Split('.'))

                            if (contains == true && sentence.Trim().Split(' ').Count() == 2)
                            {
                                //MessageBox.Show(s.Split(' ')[0].Replace(",", ""));
                                //textBox1.Text = s.Split(' ')[1] + " " + s.Split(' ')[0];
                                Int32 Avalability = CountDB("SELECT Count(*) FROM Customer WHERE Surname like '%" + SearchText.Split(' ')[0].Replace(",", "") + "%' AND FirstName like '%" + SearchText.Split(' ')[1] + "%' ORDER BY Surname ASC LIMIT 20");

                                if (Avalability == 1)
                                {
                                    try
                                    {
                                        label3.Text = "Customer Details - ";
                                       
                                        
                                    }
                                    catch { }
                                }
                                else if (Avalability == 0)
                                {
                                    label3.Text = "Customer Details";
                                    
                                }
                                cmd.CommandText = "SELECT* FROM Customer WHERE Surname like '%" + SearchText.Split(' ')[0].Replace(",", "") + "%' AND FirstName like '%" + SearchText.Split(' ')[1] + "%' ORDER BY Surname ASC LIMIT 20";
                            }
                            else
                            {

                                SearchText = SearchText.Replace(",", "");

                                foreach (var sentence1 in SearchText.TrimEnd('.').Split('.'))

                                    if (sentence1.Trim().Split(' ').Count() == 1)
                                    {
                                        Int32 Avalability = CountDB("SELECT Count(*) FROM Customer WHERE Surname like '%" + SearchText + "%' or FirstName like '%" + SearchText + "%' or UniID like '%" + SearchText + "%' or email like '%" + SearchText + "%' ORDER BY Surname ASC LIMIT 20");

                                        if (Avalability == 1)
                                        {
                                            try
                                            {
                                                label3.Text = "Customer Details - ";
                                                
                                            }
                                            catch { }
                                        }
                                        else if (Avalability == 0)
                                        {
                                            //label3.Text = "Customer Details";
                                            
                                        }
                                       
                                        cmd.CommandText = "SELECT* FROM Customer WHERE Surname like '%" + SearchText + "%' or FirstName like '%" + SearchText + "%' or UniID like '%" + SearchText + "%' or email like '%" + SearchText + "%' ORDER BY Surname ASC LIMIT 20";

                                    }
                                    else
                                    {
                                        Int32 Avalability = CountDB("SELECT Count(*) FROM Customer WHERE Surname like '%" + SearchText.Split(' ')[1] + "%' AND FirstName like '%" + SearchText.Split(' ')[0] + "%' ORDER BY Surname ASC LIMIT 20");

                                        if (Avalability == 1)
                                        {
                                            try
                                            {
                                                label3.Text = "Customer Details - ";
                                                
                                            }
                                            catch { }
                                        }
                                        else if (Avalability == 0)
                                        {
                                            //label3.Text = "Customer Details";
                                            
                                        }
                                        cmd.CommandText = "SELECT* FROM Customer WHERE Surname like '%" + SearchText.Split(' ')[1] + "%' AND FirstName like '%" + SearchText.Split(' ')[0] + "%' ORDER BY Surname ASC LIMIT 20";

                                        //MessageBox.Show(s.Split(' ')[1]);
                                    }
                            }

                        CustomerView1.DataSource = null;
                        CustomerView1.Columns.Clear();
                        CustomerView1.Rows.Clear();
                        CustomerView1.DataSource = DBGetDataSet(cmd.CommandText);

                        /*
                         using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                CustomerView1.Rows.Add(new object[]
                                {
                            reader.GetValue(0),  // U can use column index
                                    reader.GetValue(reader.GetOrdinal("ID")),  // Or column name like this
                                    reader.GetValue(reader.GetOrdinal("FirstName")),
                                    reader.GetValue(reader.GetOrdinal("Surname")),
                                    reader.GetValue(reader.GetOrdinal("Telephone")),
                                    reader.GetValue(reader.GetOrdinal("Email")),
                                    reader.GetValue(reader.GetOrdinal("UniID")),
                                    reader.GetValue(reader.GetOrdinal("Blacklisted")),
                                    reader.GetValue(reader.GetOrdinal("Mobile"))

                                });
                            }
                        }
                        */

                        //MessageBox.Show("" + CustomerView1.Rows.Count);

                        if (CustomerView1.Rows.Count == 1)
                        {
                            int FirstnameCol = FindColID(CustomerView1, "FirstName");
                            int SurnameCol = FindColID(CustomerView1, "FirstName");
                            label3.Text = "Customer Details - " + CustomerView1.CurrentRow.Cells[FirstnameCol].Value.ToString() + " " + CustomerView1.CurrentRow.Cells[SurnameCol].Value.ToString();
                            CustomerBookingsRead();
                            AvailableStockRead();
                        }
                        int BlacklistCol = FindColID(CustomerView1, "BlackListed");
                        foreach (DataGridViewRow row in CustomerView1.Rows)
                        {

                            //MessageBox.Show("" + Convert.ToDateTime(row.Cells[4].Value.ToString()));
                            
                            if (Convert.ToInt32(row.Cells[BlacklistCol].Value) == 1)
                            //dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 23:00:00
                            {
                                row.DefaultCellStyle.BackColor = Color.Orange;
                            }
                            if (Convert.ToInt32(row.Cells[BlacklistCol].Value) == 2)
                            //dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 23:00:00
                            {
                                row.DefaultCellStyle.BackColor = Color.Red;
                            }
                        }

                    }

                    Conn.Close();


                }
            }
            catch (Exception err)
            {
                HandleError("CustomersRead() : " + err);
            }
        }

        static int CountDB(string SQLString)
        {
            Int32 count = 0;

            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    
                        Conn.Open();
                        cmd.CommandText = SQLString;
                        count = Convert.ToInt32(cmd.ExecuteScalar());
                        Conn.Close();
                        //MessageBox.Show("" + count);
                    try
                    {
                    }
                    catch { }

                    


                }
            }

            return count;
        }

        static string DB_GetFirstvalue(string SQLString)
        {
            string DBValue = "";

            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {

                    Conn.Open();
                    cmd.CommandText = SQLString;
                    DBValue = Convert.ToString(cmd.ExecuteScalar());
                    //MessageBox.Show("" + cmd.ExecuteScalar());
                    Conn.Close();
                    
                    try
                    {
                    }
                    catch { }




                }
            }

            return DBValue;
        }

        static String GetNameAD(String username)
        {

            String fullName, firstName, lastName, telephone, email, LogonName;


            // connect to Active Directory
            DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://Mustang");

            DirectorySearcher searcher = new DirectorySearcher(directoryEntry);



            // apply filter to search results that will find the user
            searcher.Filter = "(&(objectClass=person) (samaccountname=" + username + "))";

            SearchResult result = searcher.FindOne();
            DirectoryEntry resultEntry = new DirectoryEntry();
            resultEntry = result.GetDirectoryEntry();

            firstName = resultEntry.Properties["cn"].Value.ToString();
            lastName = resultEntry.Properties["sn"].Value.ToString();
            telephone = resultEntry.Properties["telephoneNumber"].Value.ToString();
            email = resultEntry.Properties["mail"].Value.ToString();
            LogonName = resultEntry.Properties["samaccountname"].Value.ToString();

            //MessageBox.Show("" + LogonName);

            fullName = firstName = resultEntry.Properties["cn"].Value.ToString();



            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {

                    Conn.Open();


                    //MessageBox.Show("INSERT INTO UserData(s,e,a,n,Sol) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')");
                    cmd.CommandText = "INSERT INTO Customer (FirstName, Surname, Telephone, Email, UniID) VALUES('" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "')";
                    cmd.ExecuteNonQuery();

                    Conn.Close();


                }

            }

            return username;


        }

        void GetNameAD2()
        {
            try
            {
                if (textBox1.Text != "")
                {
                    string s = textBox1.Text;
                    bool contains = textBox1.Text.Contains(",");
                    foreach (var sentence in textBox1.Text.TrimEnd('.').Split('.'))

                        if (contains == true && sentence.Trim().Split(' ').Count() == 2)
                        {
                            //MessageBox.Show(s.Split(' ')[0].Replace(",", ""));
                            textBox1.Text = s.Split(' ')[1] + " " + s.Split(' ')[0].Replace(",", "");
                            
                        }
                        

                    //string DomainPath = "LDAP://Stirling.nene.ac.uk";
                    //string DomainPath = "LDAP://srv-adds-01.nene.ac.uk";
                    string DomainPath = "LDAP://mustang.nene.ac.uk";
                    DirectoryEntry searchRoot = new DirectoryEntry(DomainPath);
                    DirectorySearcher search = new DirectorySearcher(searchRoot);

                    search.Filter = "(&(objectClass=user) (displayname=*" + textBox1.Text + "*))";
                    //search.Filter = "(&(objectClass=person) (samaccountname=*" + textBox1.Text + "*))";

                    //search.Filter = "(&(objectClass=user)(objectCategory=person))";
                    //search.PropertiesToLoad.Add("samaccountname=mdrowla");
                    //search.PropertiesToLoad.Add("mail");
                    //search.PropertiesToLoad.Add("usergroup");
                    //search.PropertiesToLoad.Add("displayname=Mark Rowland");//first name
                    SearchResult result;
                    SearchResultCollection resultCol = search.FindAll();

                    //MessageBox.Show("" + resultCol.Count);

                    if (resultCol != null)
                    {
                        //result = resultCol[14];
                        //MessageBox.Show("" + (String)result.Properties["cn"][0] + " :");
                        //MessageBox.Show("" + (String)result.Properties["givenName"][0] + " :");
                        //MessageBox.Show("" + (String)result.Properties["telephoneNumber"][0] + " :");
                        //MessageBox.Show("" + (String)result.Properties["mail"][0] + " :");
                        //MessageBox.Show("" + (String)result.Properties["samaccountname"][0] + " :");

                        for (Int32 counter = 0; counter < resultCol.Count; counter++)
                        {

                            result = resultCol[counter];

                            //MessageBox.Show("" + (String)result.Properties["telephoneNumber"][0] + " :" + counter);
                            String firstName, lastName, telephone, email, LogonName;
                            try
                            {


                                firstName = "Unknown";
                                lastName = "Unknown";
                                telephone = "0000";
                                email = "";
                                LogonName = "Unknown";

                                    firstName = (String)result.Properties["givenName"][0];
                                    lastName = (String)result.Properties["sn"][0];
                                    LogonName = (String)result.Properties["samaccountname"][0];

                                try
                                {
                                    email = (String)result.Properties["mail"][0];
                                    telephone = (String)result.Properties["telephoneNumber"][0];
                                }
                                catch { }


                            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                                {
                                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                                    {

                                        Conn.Open();


                                        //MessageBox.Show("INSERT INTO UserData(s,e,a,n,Sol) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')");
                                        cmd.CommandText = "INSERT INTO Customer (FirstName, Surname, Telephone, Email, UniID) VALUES('" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "')";
                                        cmd.ExecuteNonQuery();

                                        Conn.Close();


                                    }

                                }
                            }
                            catch (Exception err)
                            {
                                HandleError("GetNameAD2() : " + err);
                            }

                            //*/
                        }
                    }
                }
            }

            catch (Exception err)
            {
                HandleError("GetNameAD2() : " + err);
            }


            //MessageBox.Show("Done?");

        }

        private void WritetoDB(string SQLString)
        {
            //try
            //{

            //MessageBox.Show(SQLString);
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {

                        Conn.Open();

                        cmd.CommandText = SQLString;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "update Updates Set GetUpdates = 'true by Mark' where hostname <> ''";
                        cmd.ExecuteNonQuery();


                    Conn.Close();

                    }

                }
            //}
            //catch
            //{
                //MessageBox.Show("Error");
            //}
        }

        public string DBSingleRead(string SQLString)
        {
            
                string ReturnString = "";

                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {

                        Conn.Open();
                        cmd.CommandText = SQLString;
                        ReturnString = cmd.ExecuteScalar().ToString();
                        Conn.Close();


                    }

                }
            

            return ReturnString;
        }

        private void CustomerView1_SelectionChanged(object sender, EventArgs e)
        {

            
            try
            {
                CustomerBookingsRead();

            }
            catch { }

            try
            {
                int FirstnameCol = FindColID(CustomerView1, "FirstName");
                int SurnameCol = FindColID(CustomerView1, "Surname");
                label3.Text = "Customer Details - " + CustomerView1.CurrentRow.Cells[FirstnameCol].Value.ToString() + " " + CustomerView1.CurrentRow.Cells[SurnameCol].Value.ToString();
                //AvailableStockRead();
            }
            catch { }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //Form2 op = new Form2();
            //op.Show();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            ReadTodaysBookings();
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {

        }

        private void dataGridView3_DoubleClick(object sender, EventArgs e)
        {
            
        }

        private void OutStandlingLoans_DoubleClick(object sender, EventArgs e)
        {
            

        }

        void SendConfirmation()
        {
            
            //Select DISTINCT LoanData.ID, LoanDescriptions.Description, LoanStock.AssetID, LoanData.CollectionDate, LoanData.ReturnDate from LoanData left join LoanStock on LoanData.Device = LoanStock.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID Where loandata.customer = 8 and LoanData.ConfirmationSent is null ORDER BY LoanStock.AssetID ASC
            /*
            try
            {
            //*/
                //MessageBox.Show(LoanLocationsView1.CurrentRow.Cells[1].Value.ToString());
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                    //SELECT DISTINCT LoanStock.AssetID, LoanNames.Description, Locations.LocationName FROM   LoanStock INNER JOIN LoanNames ON LoanStock.Description = LoanNames.Id INNER JOIN Locations ON LoanStock.Location = Locations.Id CROSS JOIN LoanData WHERE(NOT(LoanStock.AssetID IN (SELECT Device FROM    LoanData AS LoanData_1 WHERE(ReturnDate >= @RequestedCollectionDate) AND(CollectionDate <= @RequestedReturnDate)))) AND(LoanStock.Description = @RequestedDeviceID) AND(LoanData.Returned IS NULL)
                    int UserIDCol = FindColID(CustomerView1, "ID");
                    

                    cmd.CommandText = "Select DISTINCT LoanDescriptions.Description,  LoanStock.AssetID, LoanData.CollectionDate, LoanLocations.Location,LoanData.ReturnDate,  LoanData.ID from LoanData left join LoanStock on LoanData.Device = LoanStock.ID left join LoanLocations on LoanData.CollectionPoint = LoanLocations.ID LEFT JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID Where loandata.customer = " + CustomerView1.CurrentRow.Cells[UserIDCol].Value.ToString() + " and LoanData.ConfirmationSent is null and returned is null ORDER BY LoanStock.AssetID ASC";
                        //cmd.CommandText = "SELECT DISTINCT LoanStock.AssetID, LoanDescriptions.Description, LoanLocations.Location ,LoanStock.ID FROM LoanStock INNER JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID INNER JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID CROSS JOIN LoanData WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + CollectdateTimePicker.Value.ToString("yyyy-MM-dd") + " 13:00:00') AND (CollectionDate <= '2017-05-30 11:00:00') AND LoanData.Returned IS NULL AND LoanData.Collected IS NULL))) AND (LoanStock.Description = 1)  AND (LoanStock.Location = 1) AND LoanStock.Active = 1 ORDER BY LoanStock.ID ASC";
                        CustomerOrder.Rows.Clear();

                        using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                CustomerOrder.Rows.Add(new object[]
                                {
                                    reader.GetValue(0),  // U can use column index
                                    reader.GetValue(1),
                                    reader.GetValue(2),
                                    reader.GetValue(3),
                                    reader.GetValue(4),
                                    reader.GetValue(5)



                                });
                            }
                        }
                    
                        Conn.Close();
                    
                        if (CustomerOrder.RowCount != 0)
                        {
                        //MessageBox.Show("Here" + CustomerView1.CurrentRow.Cells[5].Value.ToString());

                       
                        int UserEmailCol = FindColID(CustomerView1, "ID");
                        int FirstnameCol = FindColID(CustomerView1, "FirstName");
                        int SurnameCol = FindColID(CustomerView1, "Surname");
                        int EmailCol = FindColID(CustomerView1, "Email");

                        if (CustomerView1.CurrentRow.Cells[5].Value.ToString() != "")
                            {
                            //MessageBox.Show("" + CustomerView1.CurrentRow.Cells[6].Value);
                            String UserEmail = CustomerView1.CurrentRow.Cells[EmailCol].Value.ToString();
                            //String FullName = CustomerView1.CurrentRow.Cells[FirstnameCol].Value.ToString() + " " + CustomerView1.CurrentRow.Cells[SurnameCol].Value.ToString();
                            String FullName = CustomerView1.CurrentRow.Cells[FirstnameCol].Value.ToString();
                            string UserID = CustomerView1.CurrentRow.Cells[UserIDCol].Value.ToString();
                                SendCalendarInvite();

                            //SendEmail(UserEmail); Replaced 18/08/17 - added UoN logo


                                SendEmailWithLogo(UserEmail, FullName);
                                
                                //DisplayMessage("Email sent to " + CustomerView1.CurrentRow.Cells[EmailCol].Value.ToString(),Color.ForestGreen);
                                //MessageBox.Show("Email sent to " + CustomerView1.CurrentRow.Cells[5].Value.ToString());
                            }
                            else
                            {
                                DisplayMessage("Email unavailable. No email address on record. " + CustomerView1.CurrentRow.Cells[EmailCol].Value.ToString(), Color.OrangeRed);
                                //MessageBox.Show("Email NOT sent. No email address on record. " + CustomerView1.CurrentRow.Cells[5].Value.ToString());
                            }
                        }

                        //Where loandata.customer = 8 and LoanData.ConfirmationSent is null AND loandata.customer = " + CustomerView1.CurrentRow.Cells[0].Value.ToString()

                        Conn.Open();
                        //MessageBox.Show("" + CustomerView1.CurrentRow.Cells[1].Value.ToString());
                        String TimeNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        cmd.CommandText = "UPDATE LoanData SET ConfirmationSent = '" + TimeNow + "', Notes = '" + NotestextBox.Text + "' Where LoanData.ConfirmationSent is null AND loandata.customer = " + CustomerView1.CurrentRow.Cells[UserIDCol].Value.ToString();
                        cmd.ExecuteNonQuery();
                        Conn.Close();

                        /*
                          
                            S (Situation): Loan Request
                            E (Escalation): 
                            A (Action): Approved
                            N (Next Step): 
                            S (Solution): Booked on Loan App

                            Notes: 

Notes: 
                         */

                        try
                        {

                            var TextString = ("S (Situation): Loan Request" + Environment.NewLine +
                                      "E (Escalation):" + Environment.NewLine +
                                      "A (Action): Approved" + Environment.NewLine +
                                      "N (Next Step):" + Environment.NewLine +
                                      "S (Solution): Booked on Loan App" + Environment.NewLine + Environment.NewLine +
                                      "Notes: ");

                            Clipboard.SetText(TextString);
                        }
                        catch { }

                    }

                    Conn.Close();


                } /*
            }
            catch (Exception err)
            {   
                HandleError("SendConfirmation() : " + err);
                //MessageBox.Show("Error..." + err);
                DisplayMessage("There was a problem. Email NOT sent. " + CustomerView1.CurrentRow.Cells[5].Value.ToString(), Color.OrangeRed );
                //MessageBox.Show("There was a problem. Email NOT sent. " + CustomerView1.CurrentRow.Cells[5].Value.ToString());


                
            }
            //*/
           
        }

        private void Confirmbutton_Click(object sender, EventArgs e)
        {
            if (CustomerView1.RowCount == 0)
            {
                DisplayMessage("No customer selected", Color.OrangeRed );
                //MessageBox.Show("No customer selected");
                return;
            }

            WaitpictureBox.Visible = true;
            Application.DoEvents();
            //ShowDialog();

            
                SendConfirmation();
            


            CustomerBookingsRead();
            AvailableStockRead();
            WaitpictureBox.Visible = false;

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //MessageBox.Show("" + e.KeyChar);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            //MessageBox.Show("");
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = false;
                WaitpictureBox.Visible = true;
                Application.DoEvents();
                //ShowDialog();
                try
                {
                    GetNameAD(textBox1.Text);
                }
                catch { }


                CustomersRead();
                CustomerBookingsRead();
                //AvailableStockRead();
                WaitpictureBox.Visible = false;
            }
        }

        private void S1button_Click(object sender, EventArgs e)
        {
            G1timer.Enabled = true;
        }

        private void G1timer_Tick(object sender, EventArgs e)
        {
            
            //SuspendLayout();
            G2timer.Enabled = false;
            G3timer.Enabled = false;

            if (G1groupBox.Height >= SlideOpen)
            {
                G1timer.Enabled = false;
                G1groupBox.Height = SlideOpen;

                G1groupBox.Top = 10;
                G2groupBox.Top = 470;
                G3groupBox.Top = 510;

            }

            else
            {

                G1groupBox.Height += SlideMove;
                G2groupBox.Top += SlideMove;
                G3groupBox.Top += SlideMove;

                /*
                if (G1groupBox.Height >= 60)
                {
                    G1groupBox.Height -= 10;
                    G2groupBox.Top -= 10;
                    G3button.Top -= 10;
                    
                }
                */

                if (G2groupBox.Height >= SlideClose)
                {
                    G2groupBox.Height -= SlideMove;
                    G3groupBox.Top -= SlideMove;
                }


                if (G3groupBox.Height >= SlideClose)
                {
                    G3groupBox.Height -= SlideMove;
                }


                G1label.Text = G1groupBox.Top.ToString();
                G2label.Text = G2groupBox.Top.ToString();
                G3label.Text = G3groupBox.Top.ToString();


            }
            
        }

        private void G2groupBox_Click(object sender, EventArgs e)
        {
            G2timer.Enabled = true;
        }

        private void G2timer_Tick(object sender, EventArgs e)
        {
            
            if (CustomerView1.Rows.Count == 0)
            {
                G1timer.Enabled = false;
                G2timer.Enabled = false;
                G3timer.Enabled = false;
                return;
            }
            int FirstnameCol = FindColID(CustomerView1, "FirstName");
            int SurnameCol = FindColID(CustomerView1, "Surname");

            label3.Text = "Customer Details - " + CustomerView1.CurrentRow.Cells[FirstnameCol].Value.ToString() + " " + CustomerView1.CurrentRow.Cells[SurnameCol].Value.ToString();
            //SuspendLayout();
            G1timer.Enabled = false;
            G3timer.Enabled = false;

            if (G2groupBox.Height >= SlideOpen)
            {
                G2timer.Enabled = false;
                G2groupBox.Height = SlideOpen;

                try
                {
                    label3.Text = "Customer Details - " + CustomerView1.CurrentRow.Cells[FirstnameCol].Value.ToString() + " " + CustomerView1.CurrentRow.Cells[SurnameCol].Value.ToString();

                    
                    //AvailableStockRead();
                }
                catch { }

            }

            else
            {

                G2groupBox.Height += SlideMove;




                if (G1groupBox.Height >= SlideClose)
                {
                    G1groupBox.Height -= SlideMove;
                    G2groupBox.Top -= SlideMove;


                }

                /*
                if (G2groupBox.Height >= 60)
                {
                    G2groupBox.Height -= 10;
                    G3button.Top -= 10;
                }
                */

                if (G3groupBox.Height >= SlideClose)
                {
                    G3groupBox.Height -= SlideMove;
                    G3groupBox.Top += SlideMove;
                }

                G1label.Text = G1groupBox.Top.ToString();
                G2label.Text = G2groupBox.Top.ToString();
                G3label.Text = G3groupBox.Top.ToString();
            }
            
        }

        private void G3timer_Tick(object sender, EventArgs e)
        {
            //SuspendLayout();
            G1timer.Enabled = false;
            G2timer.Enabled = false;

            if (G3groupBox.Height >= this.Size.Height - 150)
            {
                G3timer.Enabled = false;
                G3groupBox.Height = this.Size.Height - 150;

            }

            else
            {

                G3groupBox.Height += SlideMove;




                if (G1groupBox.Height >= SlideClose)
                {
                    G1groupBox.Height -= SlideMove;
                    G2groupBox.Top -= SlideMove;
                    G3groupBox.Top -= SlideMove;

                }


                if (G2groupBox.Height >= SlideClose)
                {
                    G2groupBox.Height -= SlideMove;
                    G3groupBox.Top -= SlideMove;
                }

                /*
                if (G3groupBox.Height >= 60)
                {
                    G3groupBox.Height -= 10;
                    G3groupBox.Top += 10;
                }
                */

                G1label.Text = G1groupBox.Top.ToString();
                G2label.Text = G2groupBox.Top.ToString();
                G3label.Text = G3groupBox.Top.ToString();
            }
            
        }

        private void G1groupBox_Click_1(object sender, EventArgs e)
        {

            G1timer.Enabled = true;

        }

        private void G2groupBox_Click_1(object sender, EventArgs e)
        {

            G2timer.Enabled = true;

        }

        private void G3groupBox_Click(object sender, EventArgs e)
        {
            G3timer.Enabled = true;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            G2timer.Enabled = true;
        }

        private void label5_Click(object sender, EventArgs e)
        {
            G3timer.Enabled = true;
        }

        private void label4_Click(object sender, EventArgs e)
        {
            G2timer.Enabled = true;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            G1timer.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (CustomerView1.RowCount == 0)
            {
                DisplayMessage("No customer selected",Color.OrangeRed );
                //MessageBox.Show("No customer selected");
                return;
            }

            if (dataGridView1.RowCount == 0)
            {
                return;
            }

            int CollectionPointIDCol = FindColID(LoanLocationsView1, "ID");
            int CustomerIDCol = FindColID(CustomerView1, "ID");
            int DeviceIDCol = FindColID(dataGridView1, "ID");

            String customerID, DeviceID, CreatedBy, CreationDate, CollectionDate, ReturnDate, Notes;
            int Assigned, CollectionPoint;
            Boolean Delivery;

            //MessageBox.Show(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            customerID = CustomerView1.CurrentRow.Cells[CustomerIDCol].Value.ToString(); //customerID
            DeviceID = dataGridView1.CurrentRow.Cells[DeviceIDCol].Value.ToString(); //DeviceID

            CreatedBy = System.Security.Principal.WindowsIdentity.GetCurrent().Name; //CreatedBy
            CreationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            CollectionPoint = Convert.ToInt32(LoanLocationsView1.CurrentRow.Cells[CollectionPointIDCol].Value);
            CollectionDate = CollectdateTimePicker.Value.ToString("yyyy-MM-dd") + " 08:30:00"; //CollectionDate
            ReturnDate = ReturndateTimePicker.Value.ToString("yyyy-MM-dd") + " 17:00:00"; //ReturnDate
            Notes = NotestextBox.Text; //Notes

            if (DeliverycheckBox.Checked == true)
            {
                Delivery = DeliverycheckBox.Checked; //Delivery
                Assigned = 1; //Assigned
            }
            else
            {
                Delivery = DeliverycheckBox.Checked; //Delivery
                Assigned = 0; //Assigned

            }


            Int32 Avalability = CountDB("SELECT COUNT(LoanStock.ID) from LoanStock WHERE LoanStock.ID = " + DeviceID + " AND (NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + ReturnDate + "') AND(CollectionDate <= '" + CollectionDate + "') AND LoanData.Returned IS NULL))) AND LoanStock.Active = 1");
            //Int32 Avalability = CountDB("SELECT COUNT(LoanStock.ID) from LoanStock WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + ReturnDate + "') AND(CollectionDate <= '" + CollectionDate + "') AND LoanData.Returned IS NULL))) AND LoanStock.Active = 1 AND LoanData.Device = " + DeviceID);

            if (Avalability == 1)
            {
                //WritetoDB("INSERT INTO LoanData( Device, Customer, Assigned, CreatedBy, CollectionDate, ReturnDate, Notes, Delivery, CollectionPoint, CreationDate ) VALUES (" + DeviceID + ", " + customerID + ", " + Assigned + ", '" + CreatedBy + "', '" + CollectionDate + "', '" + ReturnDate + "', '" + Notes + "', '" + Delivery + "', '" + CollectionPoint + "', '" + CreationDate + "')");

                if (CollectdateTimePicker.Value < DateTime.Now)
                {
                    DialogResult dialogResult = MessageBox.Show("The collection date is today. Has the item already been taken?", "Immediate Collection", MessageBoxButtons.YesNoCancel);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //Collected now...




                        WritetoDB("INSERT INTO LoanData( Device, Customer, Assigned, CreatedBy, CollectionDate, ReturnDate, Notes, Delivery, CollectionPoint, CreationDate, Collected, ReleasedBy ) VALUES (" + DeviceID + ", " + customerID + ", " + Assigned + ", '" + CreatedBy + "', '" + CollectionDate + "', '" + ReturnDate + "', '" + Notes + "', '" + Delivery + "', '" + CollectionPoint + "', '" + CreationDate + "', '" + CreationDate + "', '" + CreatedBy + "')");
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        //collecting later today
                        WritetoDB("INSERT INTO LoanData( Device, Customer, Assigned, CreatedBy, CollectionDate, ReturnDate, Notes, Delivery, CollectionPoint, CreationDate ) VALUES (" + DeviceID + ", " + customerID + ", " + Assigned + ", '" + CreatedBy + "', '" + CollectionDate + "', '" + ReturnDate + "', '" + Notes + "', '" + Delivery + "', '" + CollectionPoint + "', '" + CreationDate + "')");
                    }
                }
                else
                {
                    // not today
                    WritetoDB("INSERT INTO LoanData( Device, Customer, Assigned, CreatedBy, CollectionDate, ReturnDate, Notes, Delivery, CollectionPoint, CreationDate ) VALUES (" + DeviceID + ", " + customerID + ", " + Assigned + ", '" + CreatedBy + "', '" + CollectionDate + "', '" + ReturnDate + "', '" + Notes + "', '" + Delivery + "', '" + CollectionPoint + "', '" + CreationDate + "')");
                }

                string conStringDatosUsuarios = @"Data Source=\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                DBWright("update Updates Set GetUpdates = 'true by Mark' where hostname <> ''", conStringDatosUsuarios);


            }
            else
            {
                DisplayMessage("Stock no longer available..." + Avalability,Color.OrangeRed );
                //MessageBox.Show("Stock no longer available..." + Avalability);
            }


            CustomerBookingsRead();
            AvailableStockRead();
            ReadTodaysBookings();
        }

        private void ReturndateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            /*
            TimeSpan t = ReturndateTimePicker.Value - CollectdateTimePicker.Value;
            double NrOfDays = t.TotalDays;
            if (ReturndateTimePicker.Value.AddDays(1) < DateTime.Now)
            {
                ReturndateTimePicker.Value = DateTime.Now;
                ReturndateTimePicker.Value = ReturndateTimePicker.Value.AddDays(1);
                MessageBox.Show("Date is in the past...");

            }
            //*/

            AvailableStockRead();
        }

        private void CollectdateTimePicker_ValueChanged(object sender, EventArgs e)
        {


            TimeSpan t = ReturndateTimePicker.Value - CollectdateTimePicker.Value;
            double NrOfDays = t.TotalDays;
            if (CollectdateTimePicker.Value < DateTime.Now)
            {
                ///*
                CollectdateTimePicker.Value = DateTime.Now;
                //MessageBox.Show("Date is in the past...");
                //*/
            }

            NrOfDays = t.TotalDays;

            if (NrOfDays < 1)
            {
                ReturndateTimePicker.Value = CollectdateTimePicker.Value;
            }

            AvailableStockRead();

        }

        private void ReturndateTimePicker_Validated(object sender, EventArgs e)
        {

        }

        private void LoanDescriptionView1_SelectionChanged(object sender, EventArgs e)
        {
            AvailableStockRead();
        }

        private void LoanLocationsView1_SelectionChanged(object sender, EventArgs e)
        {
            AvailableStockRead();
        }

        private void LoanLocationsView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            AvailableStockRead();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            if (QueryAD.IsBusy != true)
            {
                QueryAD.RunWorkerAsync();
            }  
        }

        private void LoanDescriptionView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            AvailableStockRead();
        }

        private void LoanLocationsView1_Click(object sender, EventArgs e)
        {
            AvailableStockRead();
        }

        private void LoanDescriptionView1_Click(object sender, EventArgs e)
        {
            AvailableStockRead();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            WritetoDB(@"Create view Outstanding Loans as Select Firstname, Surname, LoanDescriptions.Description, ReturnDate from loandata 
                        left join LoanStock on LoanStock.ID = LoanData.Device
                        left join LoanDescriptions on LoanDescriptions.ID = LoanStock.Description
                        left join Customer on Customer.ID = LoanData.Customer
                        where returned isnull"
                    );
        }

        private void button8_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Value = dateTimePicker1.Value.AddDays(1);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Value = dateTimePicker1.Value.AddDays(-1);
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label3_TextChanged(object sender, EventArgs e)
        {
            CollectdateTimePicker.Value = DateTime.UtcNow;
            ReturndateTimePicker.Value = DateTime.UtcNow;
        }

        private void EmailConfirmlabel_Click(object sender, EventArgs e)
        {

            
        }

        void DisplayMessage(String MessageTest, Color ButtonColour)
        {
            WaitpictureBox.Invoke((MethodInvoker)delegate ()
            {
                WaitpictureBox.Visible = false;
            });

            EmailConfirmlabel.Invoke((MethodInvoker)delegate ()
             {
                 EmailConfirmlabel.BackColor = ButtonColour;
                 EmailConfirmlabel.Text = MessageTest;

                 EmailConfirmlabel.Visible = true;
                 Fadetimer.Enabled = true;
             });


            

        }

        private void Fadetimer_Tick(object sender, EventArgs e)
        {
            
                EmailConfirmlabel.Visible = false;
   
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        void HandleError(string errorcode)
        {
            WritetoDB("INSERT INTO Errors( ErrorText, Date, User) VALUES ('" + errorcode + "', '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "', '" + DateTime.Now.ToString("dd-MMM-yyyy") + "')");
        }

        private void OutStandlingLoans_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //MessageBox.Show(LoanLocationsView1.CurrentRow.Cells[1].Value.ToString());
            string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    Conn.Open();



                    string s = textBox1.Text;

                    bool contains = textBox1.Text.Contains(",");
                    foreach (var sentence in textBox1.Text.TrimEnd('.').Split('.'))

                        if (contains == true && sentence.Trim().Split(' ').Count() == 2)
                        {
                            //MessageBox.Show(s.Split(' ')[0].Replace(",", ""));
                            //textBox1.Text = s.Split(' ')[1] + " " + s.Split(' ')[0];
                            Int32 Avalability = CountDB("SELECT Count(*) FROM Customer WHERE Surname like '%" + s.Split(' ')[0].Replace(",", "") + "%' AND FirstName like '%" + s.Split(' ')[1] + "%' ORDER BY Surname ASC");

                            if (Avalability == 1)
                            {
                                try
                                {
                                    label3.Invoke((MethodInvoker)delegate () { label3.Text = "Customer Details - "; });
                                    


                                }
                                catch { }
                            }
                            else if (Avalability == 0)
                            {
                                label3.Invoke((MethodInvoker)delegate () { label3.Text = "Customer Details"; });
                                

                            }
                            cmd.CommandText = "SELECT* FROM Customer WHERE Surname like '%" + s.Split(' ')[0].Replace(",", "") + "%' AND FirstName like '%" + s.Split(' ')[1] + "%' ORDER BY Surname ASC";
                        }
                        else
                        {



                            foreach (var sentence1 in s.TrimEnd('.').Split('.'))

                                if (sentence1.Trim().Split(' ').Count() == 1)
                                {
                                    Int32 Avalability = CountDB("SELECT Count(*) FROM Customer WHERE Surname like '%" + textBox1.Text + "%' or FirstName like '%" + textBox1.Text + "%' or UniID like '%" + textBox1.Text + "%' or email like '%" + textBox1.Text + "%' ORDER BY Surname ASC");

                                    if (Avalability == 1)
                                    {
                                        try
                                        {
                                            label3.Invoke((MethodInvoker)delegate () { label3.Text = "Customer Details - "; });
                                            

                                        }
                                        catch { }
                                    }
                                    else if (Avalability == 0)
                                    {
                                        label3.Invoke((MethodInvoker)delegate () { label3.Text = "Customer Details"; });

                                    }

                                    cmd.CommandText = "SELECT* FROM Customer WHERE Surname like '%" + textBox1.Text + "%' or FirstName like '%" + textBox1.Text + "%' or UniID like '%" + textBox1.Text + "%' or email like '%" + textBox1.Text + "%' ORDER BY Surname ASC";

                                }
                                else
                                {
                                    Int32 Avalability = CountDB("SELECT Count(*) FROM Customer WHERE Surname like '%" + s.Split(' ')[1] + "%' AND FirstName like '%" + s.Split(' ')[0] + "%' ORDER BY Surname ASC");

                                    if (Avalability == 1)
                                    {
                                        try
                                        {
                                            label3.Invoke((MethodInvoker)delegate () { label3.Text = "Customer Details - "; });

                                        }
                                        catch { }
                                    }
                                    else if (Avalability == 0)
                                    {
                                        //label3.Text = "Customer Details";

                                    }
                                    cmd.CommandText = "SELECT* FROM Customer WHERE Surname like '%" + s.Split(' ')[1] + "%' AND FirstName like '%" + s.Split(' ')[0] + "%' ORDER BY Surname ASC";

                                    //MessageBox.Show(s.Split(' ')[1]);
                                }
                        }

                    //this.CustomerView1.Rows.Clear();
                    CustomerView1.Invoke((MethodInvoker)delegate () 
                    {
                        CustomerView1.DataSource = null;
                        CustomerView1.Columns.Clear();
                        CustomerView1.Rows.Clear();
                        CustomerView1.AutoGenerateColumns = true;
                        CustomerView1.DataSource = DBGetDataSet(cmd.CommandText);

                        try
                        {
                            int BlackListIDCol = FindColID(CustomerView1, "BlackListed");

                            //MessageBox.Show("" + CustomerView1.RowCount);
                            foreach (DataGridViewRow row in CustomerView1.Rows)
                            {

                                //MessageBox.Show("" + Convert.ToDateTime(row.Cells[4].Value.ToString()));

                                if (Convert.ToInt32(row.Cells[BlackListIDCol].Value) == 1)
                                {
                                    row.DefaultCellStyle.BackColor = Color.Orange;
                                }
                                if (Convert.ToInt32(row.Cells[BlackListIDCol].Value) == 2)

                                {
                                    row.DefaultCellStyle.BackColor = Color.Red;
                                }
                            }
                        }
                        catch { }

                    });



                    /*
                    using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CustomerView1.Invoke((MethodInvoker)delegate () {
                                CustomerView1.Rows.Add(new object[]
                                {
                                reader.GetValue(0),  // U can use column index
                                        reader.GetValue(reader.GetOrdinal("ID")),  // Or column name like this
                                        reader.GetValue(reader.GetOrdinal("FirstName")),
                                        reader.GetValue(reader.GetOrdinal("Surname")),
                                        reader.GetValue(reader.GetOrdinal("Telephone")),
                                        reader.GetValue(reader.GetOrdinal("Email")),
                                        reader.GetValue(reader.GetOrdinal("UniID")),
                                        reader.GetValue(reader.GetOrdinal("Blacklisted"))

                                });
                            });
                        }
                    }

                    Conn.Close();

                    */


                }




            }
        
        }

        void CustomerView1RowsClear()
        {
            if (this.CustomerView1.InvokeRequired)
            {
                CustomerView1.Invoke((MethodInvoker)delegate () { CustomerView1.Rows.Clear(); });

            }
            else
            {
                this.CustomerView1.Rows.Clear();
            }



            CustomerView1.Rows.Clear();
        }

        private void ADSearch_Tick(object sender, EventArgs e)
        {
            ADSearch.Enabled = false;
            try
            {
                if (textBox1.Text != "")
                {
                    string s = textBox1.Text;
                    bool contains = textBox1.Text.Contains(",");
                    foreach (var sentence in textBox1.Text.TrimEnd('.').Split('.'))

                        if (contains == true && sentence.Trim().Split(' ').Count() == 2)
                        {
                            textBox1.Text = s.Split(' ')[1] + " " + s.Split(' ')[0].Replace(",", "");
                        }


                    //string DomainPath = "LDAP://Stirling.nene.ac.uk";
                    //string DomainPath = "LDAP://srv-adds-01.nene.ac.uk";
                    string DomainPath = "LDAP://mustang.nene.ac.uk";

                    DirectoryEntry searchRoot = new DirectoryEntry(DomainPath);
                    DirectorySearcher search = new DirectorySearcher(searchRoot);


                    search.Filter = "(&(objectClass=user) (displayname=" + textBox1.Text + "*))";

                    SearchResult result;
                    SearchResultCollection resultCol = search.FindAll();



                    if (resultCol != null)
                    {
                        for (Int32 counter = 0; counter < resultCol.Count; counter++)
                        {

                            result = resultCol[counter];


                            String firstName, lastName, telephone, email, LogonName;


                            firstName = "Unknown";
                            lastName = "Unknown";
                            telephone = "0000";
                            email = "";
                            LogonName = "Unknown";

                            firstName = (String)result.Properties["givenName"][0];
                            lastName = (String)result.Properties["sn"][0];
                            LogonName = (String)result.Properties["samaccountname"][0];
                            //MessageBox.Show(LogonName);
                            try
                            {
                                email = (String)result.Properties["mail"][0];
                                telephone = (String)result.Properties["telephoneNumber"][0];
                            }
                            catch { }
                            Application.DoEvents();

                            try
                            {
                                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                                {
                                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                                    {

                                        Conn.Open();


                                        //MessageBox.Show("INSERT INTO UserData(s,e,a,n,Sol) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')");
                                        cmd.CommandText = "INSERT INTO Customer (FirstName, Surname, Telephone, Email, UniID) VALUES('" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "')";
                                        cmd.ExecuteNonQuery();

                                        Conn.Close();


                                    }

                                }
                            }
                            catch { }




                        }
                    }
                }







            }
            catch { }

            CustomersRead();

            WaitpictureBox.Invoke((MethodInvoker)delegate ()
            {
                WaitpictureBox.Visible = false;
            });


        }

        private void QueryAD_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                //*/
                WaitpictureBox.Invoke((MethodInvoker)delegate ()
                {
                    WaitpictureBox.Visible = true;

                });

            }
            catch { }

            //try { 
            //'P';

            //string username = "m*rowland";

            if (textBox1.Text != "")
                {

                    string UserSearchText = textBox1.Text;

                    string StudentID = textBox4.Text;
                    string DeviceID = textBox5.Text;
                    try
                    {
                        textBox4.Invoke((MethodInvoker)delegate () { textBox4.Text = ""; });
                        textBox5.Invoke((MethodInvoker)delegate () { textBox5.Text = ""; });
                    }
                    catch { }

                    if ((UserSearchText.Contains(",") == true) & (UserSearchText.Contains(", ") == false)) 
                    {
                        UserSearchText = UserSearchText.Replace(",", ", ");
                    }
               
                bool contains = UserSearchText.Contains(",");
                foreach (var sentence in UserSearchText.TrimEnd('.').Split('.'))

                    if (contains == true && sentence.Trim().Split(' ').Count() == 2)
                    {
                        UserSearchText = UserSearchText.Split(',')[1].Trim() + " " + UserSearchText.Split(',')[0].Trim();
                    }

                UserSearchText = UserSearchText.Replace(",", "");
                UserSearchText = UserSearchText.Replace(" ", "*");

                //MessageBox.Show(UserSearchText);

                var currentForest = Forest.GetCurrentForest();
                var gc = currentForest.FindGlobalCatalog();

                using (var userSearcher = gc.GetDirectorySearcher())
                {
                    userSearcher.Filter =
              "(&((&(objectCategory=Person)(objectClass=User)))(|(samAccountName=" + UserSearchText + "*)(displayname=" + UserSearchText + "*)))";
                        //(sn=" + UserSearchText + "*)
                        //(|(cn=Jim Smith)(&(givenName=Jim)(sn=Smith)))

                        userSearcher.PropertiesToLoad.Add("displayname");
                    userSearcher.PropertiesToLoad.Add("sn");
                    userSearcher.PropertiesToLoad.Add("givenName");
                    userSearcher.PropertiesToLoad.Add("samaccountname");
                    userSearcher.PropertiesToLoad.Add("mail");
                    userSearcher.PropertiesToLoad.Add("telephoneNumber");
                        userSearcher.PropertiesToLoad.Add("useraccountcontrol");
                    //userSearcher.PropertiesToLoad.Add("canonicalName");




                    SearchResultCollection resultCol2 = userSearcher.FindAll();

                        if (resultCol2 != null)
                        {
                            foreach (SearchResult userResults in resultCol2)
                            {
                                //MessageBox.Show("" + userResults.Properties["displayname"][0].ToString());
                                string firstName = "Unknown";
                                string lastName = "Unknown";
                                string telephone = "0000";
                                string email = "";
                                string LogonName = "Unknown";
                                string DomainName = "Unknown";
                                string AccountStatus = "Unknown";
                                int flags = 0;

                                try
                                {
                                    firstName = userResults.Properties["givenName"][0].ToString();
                                    lastName = userResults.Properties["sn"][0].ToString();
                                    LogonName = userResults.Properties["samaccountname"][0].ToString();
                                    DomainName = userResults.Properties["canonicalName"][0].ToString();
                                    AccountStatus = userResults.Properties["useraccountcontrol"][0].ToString();
                                    flags = (int)userResults.Properties["userAccountControl"][0];
                                }
                                catch
                                {
                                    
                                }




                                try
                                {
                                    email = userResults.Properties["mail"][0].ToString();
                                    telephone = userResults.Properties["telephoneNumber"][0].ToString();
                                }
                                catch { }

                                //MessageBox.Show(""+ Convert.ToBoolean(flags & 0x0002));
                                //MessageBox.Show("" + email);
                                if (email.Contains("@northampton.ac.uk") == true)
                                {
                                    //MessageBox.Show("Staff");
                                    //MessageBox.Show("'" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "'");
                                    WritetoDB("INSERT INTO Customer (FirstName, Surname, Telephone, Email, UniID) VALUES('" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "')");



                                }
                                else if((email.Contains("@1stdegreefacilities.co.uk") == true))
                                {
                                    WritetoDB("INSERT INTO Customer (FirstName, Surname, Telephone, Email, UniID) VALUES('" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "')");

                                }

                                else if (StudentID == LogonName || DeviceID != "")
                                {

                                    string DBStudentID = "";
                                    string DBDeviceID = "";

                                try { WritetoDB("INSERT INTO Customer (FirstName, Surname, Telephone, Email, UniID) VALUES('" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "')"); }  catch { }




                                try { DBStudentID = DBSingleRead("select ID from Customer where UniID = '"+ LogonName + "'"); } catch { DisplayMessage("Error. Student not found ID. Check ID.", Color.OrangeRed); }
                                try { DBDeviceID = DBSingleRead("select ID from LoanStock where AssetID = '"+ DeviceID.ToUpper() + "'"); } catch { DisplayMessage("Error. Device not found. Check Asset tag.", Color.OrangeRed); }




                                if (DBStudentID != "" && DBDeviceID != "") { try { StudentLoan(DBDeviceID, DBStudentID); } catch { DisplayMessage("Error. Could not reserve. ", Color.OrangeRed); } }

                                

                                }

                                else
                                {
                                    //MessageBox.Show(LogonName);
                                }


                            }

                                   
                    }


                }
            }
            /*
            }
            catch { }
            //*/

            




        }

        private void QueryAD_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            CustomersRead();

            WaitpictureBox.Invoke((MethodInvoker)delegate ()
            {
                WaitpictureBox.Visible = false;
            });
        }


        public void StudentLoan(string DeviceID, string customerID)
        {

            int Assigned, CollectionPoint;
            bool Delivery = false;
            Assigned = 0;

            string CreatedBy = System.Security.Principal.WindowsIdentity.GetCurrent().Name; //CreatedBy
            string CreationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            CollectionPoint = 5;
            string CollectionDate = DateTime.Now.ToString("yyyy-MM-dd") + " 08:30:00"; //CollectionDate
            string ReturnDate = DateTime.Now.ToString("yyyy-MM-dd") + " 17:00:00"; //ReturnDate
            string Notes = "Student Loan"; //Notes


            Int32 Avalability = CountDB("SELECT COUNT(LoanStock.ID) from LoanStock WHERE LoanStock.ID = " + DeviceID + " AND (NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + ReturnDate + "') AND(CollectionDate <= '" + CollectionDate + "') AND LoanData.Returned IS NULL))) AND LoanStock.Active = 1");
            //Int32 Avalability = CountDB("SELECT COUNT(LoanStock.ID) from LoanStock WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + ReturnDate + "') AND(CollectionDate <= '" + CollectionDate + "') AND LoanData.Returned IS NULL))) AND LoanStock.Active = 1 AND LoanData.Device = " + DeviceID);

            if (Avalability == 1)
            {
                WritetoDB("INSERT INTO LoanData( Device, Customer, Assigned, CreatedBy, CollectionDate, ReturnDate, Notes, Delivery, CollectionPoint, CreationDate, Collected, ReleasedBy ) VALUES (" + DeviceID + ", " + customerID + ", " + Assigned + ", '" + CreatedBy + "', '" + CollectionDate + "', '" + ReturnDate + "', '" + Notes + "', '" + Delivery + "', '" + CollectionPoint + "', '" + CreationDate + "', '" + CreationDate + "', '" + CreatedBy + "')");
            }
            else
            {
                DisplayMessage("Stock not longer available... Check returns.", Color.OrangeRed);
                //MessageBox.Show("Stock no longer available..." + Avalability);
            }


            textBox4.Text = "";
            textBox5.Text = "";

            CustomerBookingsRead();
            AvailableStockRead();
            ReadTodaysBookings();

        }

        private void QueryAD3_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (textBox1.Text != "")
                {
                    string s = textBox1.Text;
                    bool contains = textBox1.Text.Contains(",");
                    foreach (var sentence in textBox1.Text.TrimEnd('.').Split('.'))

                        if (contains == true && sentence.Trim().Split(' ').Count() == 2)
                        {
                            textBox1.Text = s.Split(' ')[1] + " " + s.Split(' ')[0].Replace(",", "");
                        }


                    //string DomainPath = "LDAP://Stirling.nene.ac.uk";
                    //string DomainPath = "LDAP://srv-adds-01.nene.ac.uk";
                    string DomainPath = "LDAP://mustang.nene.ac.uk";

                    DirectoryEntry searchRoot = new DirectoryEntry(DomainPath);
                    DirectorySearcher search = new DirectorySearcher(searchRoot);



                    search.Filter = "(&(objectClass=user) (displayname=*" + textBox1.Text + "*))";

                    SearchResult result;
                    SearchResultCollection resultCol = search.FindAll();



                    if (resultCol != null)
                    {
                        for (Int32 counter = 0; counter < resultCol.Count; counter++)
                        {

                            result = resultCol[counter];


                            String firstName, lastName, telephone, email, LogonName;


                            firstName = "Unknown";
                            lastName = "Unknown";
                            telephone = "0000";
                            email = "";
                            LogonName = "Unknown";

                            firstName = (String)result.Properties["givenName"][0];
                            lastName = (String)result.Properties["sn"][0];
                            LogonName = (String)result.Properties["samaccountname"][0];
                            //MessageBox.Show(LogonName);
                            try
                            {
                                email = (String)result.Properties["mail"][0];
                                telephone = (String)result.Properties["telephoneNumber"][0];
                            }
                            catch { }
                            Application.DoEvents();

                            try
                            {
                                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                                {
                                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                                    {

                                        Conn.Open();


                                        //MessageBox.Show("INSERT INTO UserData(s,e,a,n,Sol) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')");
                                        cmd.CommandText = "INSERT INTO Customer (FirstName, Surname, Telephone, Email, UniID) VALUES('" + firstName + "', '" + lastName + "', '" + telephone + "', '" + email + "', '" + LogonName + "')";
                                        cmd.ExecuteNonQuery();

                                        Conn.Close();


                                    }

                                }
                            }
                            catch { }




                        }
                    }
                }







            }
            catch { }

            

            WaitpictureBox.Invoke((MethodInvoker)delegate ()
            {
                WaitpictureBox.Visible = false;
            });

            if (ReadCustomersbackgroundWorker.IsBusy != true)
            {

                ReadCustomersbackgroundWorker.RunWorkerAsync();
            }
        }

        private void OutStandlingLoans_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {

                if (e.Button == MouseButtons.Right && OutStandlingLoans.RowCount != 0)
                {
                    var hti = OutStandlingLoans.HitTest(e.X, e.Y);
                    OutStandlingLoans.ClearSelection();
                    //OutStandlingLoans.Rows[hti.RowIndex].Selected = true;
                    //OutStandlingLoans.Rows[3].Selected = true;
                    OutStandlingLoans.CurrentCell = OutStandlingLoans.Rows[hti.RowIndex].Cells[0];
                }
            } catch { }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //CustomersRead();
            if (ReadCustomersbackgroundWorker.IsBusy != true)
            {

                ReadCustomersbackgroundWorker.RunWorkerAsync();
            }
            else
            {
                ReadCustomersbackgroundWorkerRunAgain = true;
            }



        }

        private void ReadCustomersbackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (ReadCustomersbackgroundWorkerRunAgain == true)
            {
                ReadCustomersbackgroundWorkerRunAgain = false;
                ReadCustomersbackgroundWorker.RunWorkerAsync();
            }
            else
            {
                CustomersRead();
                
            }
        }

        private void OutstandingMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            try
            {
                TimeSpan AvailbleDays = TimeSpan.FromDays(14);
                DateTime ReturnDateAsDate = Convert.ToDateTime(OutStandlingLoans.CurrentRow.Cells[5].Value);
                DateTime CollectionDateAsDate = Convert.ToDateTime(OutStandlingLoans.CurrentRow.Cells[4].Value);
                string ReturnDate = ReturnDateAsDate.ToString("yyyy-MM-dd") + " 10:00:00";
                string CollectionDate = CollectionDateAsDate.ToString("yyyy-MM-dd") + " 14:00:00";

                

                int DeviceType = Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[8].Value);
                int Location = Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[10].Value);
                string CollectedDate = "";
                try
                {
                    CollectedDate = OutStandlingLoans.CurrentRow.Cells[6].Value.ToString();
                    
                }
                catch
                {
                    
                }

                if (CollectedDate != "") { toolStripTextBoxCancel.Visible = false; }
                

                //MessageBox.Show("DeviceType=" + DeviceType);

                toolStripComboBox1.Items.Clear();
                toolStripComboBox2.Items.Clear();

                toolStripComboBox1.Text = "";
                toolStripComboBox2.Text = "";


                try
                {
                    string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                    using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                    {
                        using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                        {

                            Conn.Open();



                            String DeviceID = OutStandlingLoans.CurrentRow.Cells[9].Value.ToString();
                            String DateTimeString = ReturnDateAsDate.ToString("yyyy-MM-dd") + " 08:31:00"; //ReturnDate


                            ReturnDateAsDate = Convert.ToDateTime(DateTimeString);


                            //MessageBox.Show("DeviceID=" + OutStandlingLoans.CurrentRow.Cells[9].Value.ToString());

                            //MessageBox.Show("INSERT INTO UserData(s,e,a,n,Sol) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')");
                            cmd.CommandText = "select * from loandata where device = '" + DeviceID + "' and collectiondate > '" + DateTimeString + "' and returned is null order by collectiondate desc";

                            //MessageBox.Show("Here");

                            using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                            {


                                while (reader.Read())
                                {




                                    AvailbleDays = Convert.ToDateTime(reader.GetValue(reader.GetOrdinal("collectiondate"))) - ReturnDateAsDate;


                                    //MessageBox.Show("Availble Days=" + AvailbleDays);




                                }


                            }

                            cmd.CommandText = "SELECT DISTINCT LoanStock.AssetID, LoanDescriptions.Description, LoanLocations.Location ,LoanStock.ID FROM LoanStock INNER JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID INNER JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID CROSS JOIN LoanData WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE((ReturnDate >= '" + CollectionDate  + "') AND (CollectionDate <= '" + ReturnDate + "')) AND LoanData.Returned IS NULL))) AND (LoanStock.Description = " + DeviceType + ")  AND (LoanStock.Location = " + Location + ") AND LoanStock.Active = 1 ORDER BY LoanStock.ID ASC";
                            //cmd.CommandText = "SELECT DISTINCT LoanStock.AssetID, LoanDescriptions.Description, LoanLocations.Location ,LoanStock.ID FROM LoanStock INNER JOIN LoanDescriptions ON LoanStock.Description = LoanDescriptions.ID INNER JOIN LoanLocations ON LoanStock.Location = LoanLocations.ID CROSS JOIN LoanData WHERE(NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + CollectdateTimePicker.Value.ToString("yyyy-MM-dd") + " 13:00:00') AND (CollectionDate <= '2017-05-30 11:00:00') AND LoanData.Returned IS NULL AND LoanData.Collected IS NULL))) AND (LoanStock.Description = 1)  AND (LoanStock.Location = 1) AND LoanStock.Active = 1 ORDER BY LoanStock.ID ASC";

                            //MessageBox.Show("Here");

                            using (System.Data.SQLite.SQLiteDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {

                                    toolStripComboBox2.Items.Add(reader.GetValue(reader.GetOrdinal("AssetID")));


                                }
                            }

                            Conn.Close();

                            if (toolStripComboBox2.Items.Count == 0)
                            {
                                toolStripComboBox2.Visible = false;
                                confirmToolStripMenuItem.Visible = false;
                                toolStripMenuItem3.Visible = true;
                            }
                            else
                            {
                                toolStripComboBox2.Visible = true;
                                confirmToolStripMenuItem.Visible = true;
                                toolStripMenuItem3.Visible = false;

                            }


                        }

                    }
                }
                catch { }



                int Counter = 14;

                if (Convert.ToInt32(AvailbleDays.Days) < Counter)
                {
                    Counter = Convert.ToInt32(AvailbleDays.Days);
                }

                //MessageBox.Show("" + Counter);

                if (Counter <= 1)
                {
                    toolStripComboBox1.Visible = false;
                    ConfirmChangetoolStripMenuItem1.Visible = false;
                    toolStripMenuItem1.Visible = true;
                }
                else
                {
                    toolStripComboBox1.Visible = true;
                    ConfirmChangetoolStripMenuItem1.Visible = true;
                    toolStripMenuItem1.Visible = false;
                }



                for (int i = 1; i < Counter; i++)
                {
                    //MessageBox.Show("" + Counter);
                    toolStripComboBox1.Items.Add("+" + i + " (" + ReturnDateAsDate.AddDays(i).ToString("dd-MMM") + " - " + ReturnDateAsDate.AddDays(i).DayOfWeek.ToString() +  ")");
                }
            } catch { }
        

            
        }

        private void ConfirmChangetoolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {


                //MessageBox.Show("" + OutStandlingLoans.CurrentRow.Cells[0].Value.ToString());

                DateTime NewReturnDate = Convert.ToDateTime(OutStandlingLoans.CurrentRow.Cells[5].Value).AddDays(toolStripComboBox1.SelectedIndex + 1);
                //MessageBox.Show("" + NewReturnDate);

                if (toolStripComboBox1.Text == "") { return; }
                WritetoDB("UPDATE LoanData SET ReturnDate = '" + NewReturnDate.ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE ID = " + OutStandlingLoans.CurrentRow.Cells[0].Value.ToString() + "");
                WritetoDB("Insert into Ledger (User, Action, EffectedTable, EffectedID, Date) VALUES ('"+ System.Security.Principal.WindowsIdentity.GetCurrent().Name + "','Extend Booking from "+ OutStandlingLoans.CurrentRow.Cells[5].Value + " to "+ NewReturnDate.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'UPDATE LoanData', '" + OutStandlingLoans.CurrentRow.Cells[0].Value.ToString() + "', '" + NewReturnDate.ToString("yyyy-MM-dd HH:mm:ss") + "')");
                CustomerBookingsRead();
                DisplayMessage("Booking extended until " + NewReturnDate.ToString("dd-MMM-yyyy HH:mm"),Color.ForestGreen);
            } catch { }

        }

        private void CustomerView1_SelectionChanged_1(object sender, EventArgs e)
        {
            NotestextBox.Text = "";

        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == (char)13)
            {
                CustomerBookingsRead();
                G2timer.Enabled = true;
            }
        }

        private void textBox1_KeyDown_1(object sender, KeyEventArgs e)
        {
            try
            {


                //MessageBox.Show("" + e.KeyCode);
                if (e.KeyCode == Keys.Down)
                {
                    //MessageBox.Show("" + CustomerView1.CurrentCell.RowIndex);
                    int NewRow = CustomerView1.CurrentCell.RowIndex + 1;
                    CustomerView1.CurrentCell = CustomerView1.Rows[NewRow].Cells[2];
                }

                if (e.KeyCode == Keys.Up)
                {
                    int NewRow = CustomerView1.CurrentCell.RowIndex - 1;
                    CustomerView1.CurrentCell = CustomerView1.Rows[NewRow].Cells[2];
                }

                if (e.KeyCode == Keys.Return)
                {
                    /*
                    if (QueryAD.IsBusy != true)
                    {
                        QueryAD.RunWorkerAsync();
                    }
                    */
                }


            }
            catch { }
        }

        private void confirmToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                int DeviceID = Convert.ToInt32(DB_GetFirstvalue("select id from loanstock where AssetID = '" + toolStripComboBox2.Text + "'"));
                int LoanID = Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[0].Value);

                DateTime ReturnDateAsDate = Convert.ToDateTime(OutStandlingLoans.CurrentRow.Cells[5].Value);
                DateTime CollectionDateAsDate = Convert.ToDateTime(OutStandlingLoans.CurrentRow.Cells[4].Value);

                string CollectionDate = ReturnDateAsDate.ToString("yyyy-MM-dd") + " 08:30:00";
                string ReturnDate = CollectionDateAsDate.ToString("yyyy-MM-dd") + " 17:00:00";





                Int32 Avalability = CountDB("SELECT COUNT(LoanStock.ID) from LoanStock WHERE LoanStock.ID = " + DeviceID + " AND (NOT(LoanStock.ID IN (SELECT Device FROM LoanData WHERE(ReturnDate >= '" + ReturnDate + "') AND(CollectionDate <= '" + CollectionDate + "') AND LoanData.Returned IS NULL))) AND LoanStock.Active = 1");
                //MessageBox.Show("" + OutStandlingLoans.CurrentRow.Cells[0].Value + " - " + DeviceID);
                if (Avalability == 1)
                {
                    
                    WritetoDB("update LoanData set Device = " + DeviceID + " where id = " + LoanID);
                    DisplayMessage("Device changed to " + toolStripComboBox2.Text,Color.ForestGreen );

                }
                else
                {
                    DisplayMessage("Change failed!! Stock no longer available.",Color.OrangeRed );
                }

                CustomerBookingsRead();
            }
            catch { }   

            //update LoanData set Device = 14 where id = 286
        }

        private void goToCustomerToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                goToCustomer_toolStripMenuItem__dataGridView2.Visible = true;
                goToCustomer_toolStripMenuItem__dataGridView3.Visible = false;


                if (e.Button == MouseButtons.Right && dataGridView2.RowCount != 0)
                {
                    var hti = dataGridView2.HitTest(e.X, e.Y);
                    dataGridView2.ClearSelection();
                    //OutStandlingLoans.Rows[hti.RowIndex].Selected = true;
                    //OutStandlingLoans.Rows[3].Selected = true;
                    dataGridView2.CurrentCell = dataGridView2.Rows[hti.RowIndex].Cells[0];
                }
            } catch { }
    
        }

        private void dataGridView3_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                goToCustomer_toolStripMenuItem__dataGridView2.Visible = false;
                goToCustomer_toolStripMenuItem__dataGridView3.Visible = true;

                if (e.Button == MouseButtons.Right && dataGridView3.RowCount != 0)
                {
                    var hti = dataGridView3.HitTest(e.X, e.Y);
                    dataGridView3.ClearSelection();
                    //OutStandlingLoans.Rows[hti.RowIndex].Selected = true;
                    //OutStandlingLoans.Rows[3].Selected = true;
                    dataGridView3.CurrentCell = dataGridView3.Rows[hti.RowIndex].Cells[0];
                }
            } catch { }
        }

        private void UserLoansView_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {

                if (e.Button == MouseButtons.Right && UserLoansView.RowCount != 0)
                {
                    var hti = UserLoansView.HitTest(e.X, e.Y);
                    UserLoansView.ClearSelection();
                    //OutStandlingLoans.Rows[hti.RowIndex].Selected = true;
                    //OutStandlingLoans.Rows[3].Selected = true;
                    UserLoansView.CurrentCell = UserLoansView.Rows[hti.RowIndex].Cells[0];
                }
            } catch { }
        }

        private void goToCustomer_toolStripMenuItem__dataGridView2_Click(object sender, EventArgs e)
        {
            try
            {
                textBox1.Text = "" + dataGridView2.CurrentRow.Cells[0].Value;

                G1timer.Enabled = true;
            } catch { }
        }

        private void goToCustomer_toolStripMenuItem__dataGridView3_Click(object sender, EventArgs e)
        {
            try
            {
                textBox1.Text = "" + dataGridView3.CurrentRow.Cells[0].Value;

                G1timer.Enabled = true;
            } catch { }
        }

        private void LoadUserMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            //MessageBox.Show("" + sender);
        }

        private void G2label_Click(object sender, EventArgs e)
        {

        }

        private void CustomerView1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {

                if (e.Button == MouseButtons.Right && CustomerView1.RowCount != 0)
                {
                    var hti = CustomerView1.HitTest(e.X, e.Y);
                    CustomerView1.ClearSelection();
                    //OutStandlingLoans.Rows[hti.RowIndex].Selected = true;
                    //OutStandlingLoans.Rows[3].Selected = true;
                    CustomerView1.CurrentCell = CustomerView1.Rows[hti.RowIndex].Cells[0];
                }
            }
            catch { }
        }

        private void CustomerView1_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                int FirstnameCol = FindColID(CustomerView1, "FirstName");
                int SurnameCol = FindColID(CustomerView1, "Surname");
                CustomerBookingsRead();
                label3.Text = "Customer Details - " + CustomerView1.CurrentRow.Cells[FirstnameCol].Value.ToString() + " " + CustomerView1.CurrentRow.Cells[SurnameCol].Value.ToString();
                LoanLocationsView1.ClearSelection();
                LoanLocationsView1.CurrentCell = LoanLocationsView1.Rows[0].Cells[2];

            }
            catch { }
        }

        private void NotestextBox_Enter(object sender, EventArgs e)
        {
            //MessageBox.Show("" + NotestextBox.Width);

            this.NotestextBox.Location = new Point(271, 189);

            NotestextBox.Width = 364;
            NotestextBox.Height = 241;
            Minimizebutton.Visible = true;


        }

        private void NotestextBox_Leave(object sender, EventArgs e)
        {
            
            try
            {
                string value = NotestextBox.Text;

                int PosA = value.IndexOf("collection date") + 20;
                int PosB = value.IndexOf("Please specify loan return date") - 1;
                int PosC = value.IndexOf("return date") + 16;
                int PosD = value.IndexOf("Select Location for Collection/Return") - 1;

                String StringCollectionDate = value.Substring(PosA, PosB - PosA);
                String StringReturnDate = value.Substring(PosC, PosD - PosC);

                //MessageBox.Show("" + StringCollectionDate + " - " + StringReturnDate);
                CollectdateTimePicker.Value = Convert.ToDateTime(StringCollectionDate);
                ReturndateTimePicker.Value = Convert.ToDateTime(StringReturnDate);

            }
            catch
            {
                DisplayMessage("Dates NOT captured. Please check", Color.DarkOrange);
            }

            try
            {
                string value = NotestextBox.Text;
                string StringCampus = value.Substring(value.IndexOf("Collection/Return") + 22, 4);

                //MessageBox.Show("" + StringCampus);

                if (StringCampus == "Park")
                {
                    LoanLocationsView1.CurrentCell = LoanLocationsView1.Rows[1].Cells[2];
                }

                if (StringCampus == "Aven")
                {
                    LoanLocationsView1.CurrentCell = LoanLocationsView1.Rows[2].Cells[2];
                }

            }
            catch
            {
                DisplayMessage("Campus NOT captured. Please check", Color.DarkOrange);
            }


            this.NotestextBox.Location = new Point(571, 389);

            NotestextBox.Width = 64;
            NotestextBox.Height = 41;
            Minimizebutton.Visible = false;


        }

        private void tabControl2_DrawItem(object sender, DrawItemEventArgs e)
        {

        }

        private void TabFlashtimer_Tick(object sender, EventArgs e)
        {
            if (currentColor == Color.Yellow)
            {
                currentColor = Color.Green;
                tabControl2.TabPages[2].Text = "ALERT! ATTENTION REQUIRED";
            }
            else
            {


                currentColor = Color.Yellow;
                tabControl2.TabPages[2].Text = "                                  ";
            }
            //tabControl2.Refresh();
        }

        private void dataGridView2_DoubleClick_1(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {




            if (Convert.ToInt32(dataGridView2.CurrentRow.Cells[6].Value) != 1)
            {
                //MessageBox.Show("" + System.Security.Principal.WindowsIdentity.GetCurrent().Name);

                String UserName = dataGridView2.CurrentRow.Cells[0].Value.ToString();
                String Device = dataGridView2.CurrentRow.Cells[1].Value.ToString();
                String AssetTag = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                Int32 UserID = Convert.ToInt32(dataGridView2.CurrentRow.Cells[6].Value);

                String MSGMessage = "Has " + UserName + " collected the " + Device + " (" + AssetTag + ")?";
                String MSGTitle = "Confirm release";

                DialogResult dialogResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    WritetoDB("UPDATE LoanData SET Collected = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', ReleasedBy = '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' WHERE Id = " + UserID);
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }

            }

            ReadTodaysBookings();
        }

        private void dataGridView3_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Convert.ToInt32(dataGridView3.CurrentRow.Cells[7].Value) != 1)
            {
                

                String UserName = dataGridView3.CurrentRow.Cells[0].Value.ToString();
                String Device = dataGridView3.CurrentRow.Cells[2].Value.ToString();
                String AssetTag = dataGridView3.CurrentRow.Cells[3].Value.ToString();
                Int32 UserID = Convert.ToInt32(dataGridView3.CurrentRow.Cells[7].Value);

                String MSGMessage = "Has " + UserName + " returned the " + Device + " (" + AssetTag + ")?";
                String MSGTitle = "Confirm return";

                DialogResult dialogResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    WritetoDB("UPDATE LoanData SET Returned = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', ReturnedBy = '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' WHERE Id = " + UserID);
                    

                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }

            }

            ReadTodaysBookings();
        }

        void OutStandlingLoans_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                string HasBeenCollectedDate = "" + OutStandlingLoans.CurrentRow.Cells[6].Value.ToString();
                Int32 LoanID = Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[0].Value);
                string UserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                string ThisTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                //MessageBox.Show("" + OutStandlingLoans.CurrentRow.Cells[6].Value.ToString());




                //UPDATE LoanData SET Collected = 1 WHERE Id = 52;
                string conStringDatosUsuarios = @"\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\LoanAppData.db3";
                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();

                        

                        if (HasBeenCollectedDate != "")
                        {

                            String MSGMessage = "Do you want to return this device?";
                            String MSGTitle = "Return device";
                            



                            DialogResult dialogResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
                            if (dialogResult == DialogResult.Yes)
                            {
                                WritetoDB("UPDATE LoanData SET Returned = '" + ThisTime + "', ReturnedBy = '" + UserName + "' WHERE Id = " + LoanID);
                                //cmd.CommandText = "UPDATE LoanData SET Returned = '" + ThisTime + "' WHERE Id = " + Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[0].Value);
                                //cmd.ExecuteNonQuery();
                            }
                            else if (dialogResult == DialogResult.No)
                            {
                                //do something else
                            }
                        }

                        
                        
                        //MessageBox.Show("" + Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[0].Value));

                        if (HasBeenCollectedDate == "")
                        {

                            string MSGMessage = "Do you want to cancel this loan request?";
                            string MSGTitle = "Cancel loan.";

                            DialogResult CancelResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
                            if (CancelResult == DialogResult.Yes)
                            {
                                WritetoDB("UPDATE LoanData SET Collected = '" + ThisTime + "', Returned = '" + ThisTime + "', ReturnedBy = '" + UserName + "' WHERE Id = " + LoanID);

                                //cmd.CommandText = "UPDATE LoanData SET Collected = '" + ThisTime + "' WHERE Id = " + Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[0].Value);
                                //cmd.ExecuteNonQuery();
                            }
                            else if (CancelResult == DialogResult.No)
                            {
                                //do something else
                            }



                        }



                        Conn.Close();
                    }
                }


                CustomerBookingsRead();
                //AvailableStockRead();
                ReadTodaysBookings();
            }
            catch { }
        }

        private void CustomerView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try {
                int UserIDCol = FindColID(CustomerView1, "ID");
                int UniIDCol = FindColID(CustomerView1, "UniID");
                Int32 UserID = Convert.ToInt32(CustomerView1.CurrentRow.Cells[UserIDCol].Value);
                string UserName = CustomerView1.CurrentRow.Cells[UniIDCol].Value.ToString();
                string CurrentMobile = "";




                if (CustomerView1.CurrentRow.Cells[8].Value != null)
                {
                    int MobileCol = FindColID(CustomerView1, "Mobile");
                    CurrentMobile = CustomerView1.CurrentRow.Cells[MobileCol].Value.ToString();
                }


                //MessageBox.Show("" + UserID);

                String Response = Microsoft.VisualBasic.Interaction.InputBox("Users Mobile number?", "Add / Change Mobile number for " + UserName, CurrentMobile);



                if (Response != "")
                {
                    //MessageBox.Show(Response + "");
                    WritetoDB("UPDATE Customer SET Mobile = '" + Response + "' WHERE ID = " + UserID);
                    CustomersRead();
                }

            }
            catch { }
            }

        private void toolStripTextBoxCancel_Click(object sender, EventArgs e)
        {

            string HasBeenCollectedDate = "" + OutStandlingLoans.CurrentRow.Cells[6].Value.ToString();
            Int32 LoanID = Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[0].Value);
            string UserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string ThisTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            string MSGMessage = "Do you want to cancel this loan request?";
            string MSGTitle = "Cancel loan.";

            DialogResult CancelResult = MessageBox.Show(MSGMessage, MSGTitle, MessageBoxButtons.YesNo);
            if (CancelResult == DialogResult.Yes)
            {
                WritetoDB("UPDATE LoanData SET Collected = '" + ThisTime + "', Returned = '" + ThisTime + "', ReturnedBy = '" + UserName + "' WHERE Id = " + LoanID);

                //cmd.CommandText = "UPDATE LoanData SET Collected = '" + ThisTime + "' WHERE Id = " + Convert.ToInt32(OutStandlingLoans.CurrentRow.Cells[0].Value);
                //cmd.ExecuteNonQuery();
            }
            else if (CancelResult == DialogResult.No)
            {
                //do something else
            }

            CustomerBookingsRead();

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {


            
                if (textBox3.TextLength == 8)
            {
                if (textBox3.Text.Substring(0, 3) == "UON" || textBox3.Text.Substring(0, 3) == "uon")
                {
                    textBox5.Text = textBox3.Text;
                    textBox3.Text = "";
                }
                else
                {
                    textBox4.Text = textBox3.Text;
                    textBox3.Text = "";
                }
            }

            if (textBox4.TextLength != 0 && textBox5.TextLength != 0)
            {
                textBox1.Text = textBox4.Text;
                if (QueryAD.IsBusy != true)
                {
                    QueryAD.RunWorkerAsync();
                }
            }
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && dataGridView1.RowCount != 0)
            {
                var hti = dataGridView1.HitTest(e.X, e.Y);
                dataGridView1.ClearSelection();
                //OutStandlingLoans.Rows[hti.RowIndex].Selected = true;
                //OutStandlingLoans.Rows[3].Selected = true;
                dataGridView1.CurrentCell = dataGridView1.Rows[hti.RowIndex].Cells[0];
            }
        }

        private void CustomerView1_DataSourceChanged(object sender, EventArgs e)
        {
            
        }

        private void contextMenuStripMIA_Opening(object sender, CancelEventArgs e)
        {

        }

        private void itemIsMissingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int DeviceIDCol = FindColID(dataGridView1, "ID");

            
            int customerID = 1;
            int Assigned = 1;
            int CollectionPoint = 0;
           
            string DeviceID = dataGridView1.CurrentRow.Cells[DeviceIDCol].Value.ToString(); //DeviceID
            string CreatedBy = System.Security.Principal.WindowsIdentity.GetCurrent().Name; //CreatedBy
            string CreationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string CollectionDate = DateTime.Now.ToString("yyyy-MM-dd" + " 08:00:00");
            string ReturnDate = "2030-01-01 16:30:00";
            string Notes = "Missing in action";
            string Delivery = "false";

            WritetoDB("INSERT INTO LoanData( Device, Customer, Assigned, CreatedBy, CollectionDate, ReturnDate, Notes, Delivery, CollectionPoint, CreationDate, Collected, ReleasedBy ) VALUES (" + DeviceID + ", " + customerID + ", " + Assigned + ", '" + CreatedBy + "', '" + CollectionDate + "', '" + ReturnDate + "', '" + Notes + "', '" + Delivery + "', '" + CollectionPoint + "', '" + CreationDate + "', '" + CreationDate + "', '" + CreatedBy + "')");

            AvailableStockRead();

        }
    }
    
    
}



            
       

