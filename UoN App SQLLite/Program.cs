using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using UoN_App_SQLLite;
    

namespace UoN_App_SQLLite
{

    public class MultiFormContext : ApplicationContext
    {
        private int openForms;
        public MultiFormContext(params Form[] forms)
        {
            openForms = forms.Length;

            foreach (var form in forms)
            {
                form.FormClosed += (s, args) =>
                {
                    //When we have closed the last of the "starting" forms, 
                    //end the program.
                    if (System.Threading.Interlocked.Decrement(ref openForms) == 0)
                        ExitThread();
                };

                form.Show();
            }
        }
    }

    static class Program
    {
        

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //try {

                string subPath = @"\UoN App"; // your code goes here
                                              //MessageBox.Show(subPath);
                bool DBFile = System.IO.File.Exists(@"\UoN App\UoNSeansData.db3");

                bool exists = System.IO.Directory.Exists(subPath);


                if (!exists)
                {

                    //MessageBox.Show("No Folder");
                    System.IO.Directory.CreateDirectory(subPath);
                }
                string conStringDatosUsuarios = @"\UoN App\UoNSeansData.db3";
                if (!DBFile)
                {
                    //MessageBox.Show("No Database");

                    System.Data.SQLite.SQLiteConnection.CreateFile(conStringDatosUsuarios);

                }

                //string conStringDatosUsuarios = @"\UoN App\UoNSeansData.db3";
                string createQuery = @"CREATE TABLE IF NOT EXISTS
                                    [UserData] (
                                    [Id] INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,  
                                    [s] NVARCHAR(254) NULL,
                                    [e] NVARCHAR(254) NULL,
                                    [a] NVARCHAR(254) NULL,
                                    [n] NVARCHAR(254) NULL,
                                    [Sol] NVARCHAR(254) NULL
                                    )";

                using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection("data source=" + conStringDatosUsuarios))
                {
                    using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                    {
                        Conn.Open();
                        cmd.CommandText = createQuery;
                        cmd.ExecuteNonQuery();
                        //cmd.CommandText = "INSERT INTO UserData(s,e,a,n,Sol) values('Test1','Test2','Test3','Test4','Test5')";
                        //cmd.();

                        cmd.CommandText = "SELECT * FROM UserData";



                        Conn.Close();



                    }
                }



                Application.EnableVisualStyles();

                Application.SetCompatibleTextRenderingDefault(true);
                //Application.Run(new Form1());

                Application.Run(new MultiFormContext(new FormSEANS()));
            

            /*
            {
            catch
            {
                MessageBox.Show("There was a fatal error... Check C:\\UoN App\\ Rename folder and try again.");
            }
            */
        }

    }
}
