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
    public partial class FormDataRecovery : Form
    {
        public FormDataRecovery()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                string S = Application.StartupPath;
                int idx = S.LastIndexOf((char)92);
                string SearchPath = S.Substring(0, idx);

                //MessageBox.Show(SearchPath);

                foreach (string file in System.IO.Directory.EnumerateFiles(SearchPath,
                            "*.db3",
                            System.IO.SearchOption.AllDirectories))
                {
                    // Display file path.
                    //MessageBox.Show(file);

                    DateTime lastModified = System.IO.File.GetLastWriteTime(file);

                    System.IO.File.Move(file, @"C:\UoN App\UoNSeansData " + " " + lastModified.ToString("dd-MM-yy HH-mm-ss") + ".db3");

                    //dataGridViewDBFiles.CurrentRow.Cells[0].Value = file;



                    //dataGridViewDBFiles.CurrentRow.Cells[1].Value = (lastModified.ToString("dd-MM-yy HH-mm-ss"));

                }


                string UserDatePath = @"C:\UoN App";
                //MessageBox.Show(UserDatePath);

                foreach (string Localfile in System.IO.Directory.EnumerateFiles(UserDatePath,
                            "*.db3",
                            System.IO.SearchOption.AllDirectories))
                {
                    // Display file path.
                    //MessageBox.Show(Localfile);

                    DateTime lastModified = System.IO.File.GetCreationTime(Localfile);

                    this.dataGridViewDBFiles.Rows.Add(Localfile, lastModified.ToString("dd-MM-yy HH-mm-ss"));

                    //dataGridViewDBFiles.CurrentRow.Cells[0].Value = Localfile;
                    //dataGridViewDBFiles.CurrentRow.Cells[1].Value = (lastModified.ToString("dd-MM-yy HH-mm-ss"));

                    


                }



            }
            catch { }
        }

        private void dataGridViewDBFiles_DoubleClick(object sender, EventArgs e)
        {
            MessageBox.Show(dataGridViewDBFiles.CurrentRow.Cells[0].Value.ToString());

            String CurrentFile = @"C:\UoN App\UoNSeansData.db3";
            DateTime lastModified = System.IO.File.GetCreationTime(CurrentFile);
            String CurrentFileNewName = CurrentFile + " " + lastModified.ToString("dd-MM-yy HH-mm-ss" + ".db3");

            String RestoreFileName = dataGridViewDBFiles.CurrentRow.Cells[0].Value.ToString();

           
            MessageBox.Show(CurrentFile + "" + RestoreFileName);
            if (CurrentFile != RestoreFileName)
            {

            System.IO.File.Move(CurrentFile, CurrentFileNewName);
            System.IO.File.Copy(RestoreFileName, CurrentFile);

            }
            
            
            
        }
    }
}
