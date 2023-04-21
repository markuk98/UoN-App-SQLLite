using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SQLite;

namespace UoN_App_SQLLite
{
    class ROSDBQ
    {
        string Connection = @"Data Source=\\\\\northampton\shared\INS\IS_Info\A_Service Strategy\Teams\Service Delivery\Audio Visual and Events\Loan Database\ROS.db3";

        public DataTable DBGetDataSet(string SQL)
        {
            // Requires using System.Data;

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

        public void DBWrite(string SQLString)
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

        public string DBSingleRead(string SQLString)
        {
            string ReturnString = "";

            using (System.Data.SQLite.SQLiteConnection Conn = new System.Data.SQLite.SQLiteConnection(Connection))
            {
                using (System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand(Conn))
                {
                    try
                    {
                        Conn.Open();
                        cmd.CommandText = SQLString;
                        ReturnString = cmd.ExecuteScalar().ToString();
                        Conn.Close();
                    }
                    catch { }
                }
            }

            return ReturnString;
        }
    }
}

