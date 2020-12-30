using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailGonderme
{
    class Veri
    {
        public static DataTable DtAl(string sqlSorgu)
        {
            int sqltimeout = int.Parse(ConfigurationManager.AppSettings["sqlTimeout"].ToString());

            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(ConfigurationManager.AppSettings["sqlBaglanti"]))
            {
                try
                {
                    conn.Open();
                    SqlDataAdapter adpt = new SqlDataAdapter(sqlSorgu, conn);
                    adpt.SelectCommand.CommandTimeout = sqltimeout;
                    adpt.Fill(dt);
                }
                catch (Exception ex)
                {
                    conn.Close();
                   
                    throw new Exception("Datatable alınamadı !! " + ex.Message);
                }
            }

            return dt;
        }
    }
}
