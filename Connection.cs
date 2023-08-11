using Humanizer;
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
        public string connectionString = @"Data Source=Server,Port;Initial Catalog=BDCliente;User ID=user;Password=123456";

        public static SqlCommand CreateCommand(string queryString)
        {
            SqlCommand command = new SqlCommand();
            string connectionString = @"Data Source=Server,Port;Initial Catalog=BDCliente;User ID=user;Password=123456";

            using (SqlConnection connection = new SqlConnection(
                       connectionString))
            {
                if (queryString != null) 
                { 
                    //MessageBox.Show("Informe uma Consulta SQL", "Erro!", MessageBoxButtons.OKCancel);
                    return command;
                }

                command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
                return command;
            }



        }
    }
}
