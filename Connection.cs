using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InteropWord
{
    public class Connection
    {
        public string connectionString = @"Data Source=152.67.36.137,49205;Initial Catalog=RIOPLASTIC-HML;User ID=desenv;Password=crhumanos321";

        public static SqlCommand CreateCommand(string queryString)
        {
            string connectionString = @"Data Source=152.67.36.137,49205;Initial Catalog=RIOPLASTIC-HML;User ID=desenv;Password=crhumanos321";

            using (SqlConnection connection = new SqlConnection(
                       connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
                return command;
            }
        }
    }
}
