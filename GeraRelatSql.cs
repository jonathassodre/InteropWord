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
using static InteropWord.Connection;

namespace InteropWord
{
    public partial class GeraRelatSql : Form
    {

        Connection conn = new Connection();
        string conteudoArquivo = string.Empty;
        string caminhoArquivo = string.Empty;
        string nomeArquivo = string.Empty;
        private Dictionary<string, string> tabelaComNomesEValores = new Dictionary<string, string>();


        public GeraRelatSql()
        {
            InitializeComponent();
        }

        string query = "";
        DataTable dataTable = new DataTable();

        private void btnTestar_Click(object sender, EventArgs e)
        {
            try
            {
                query = txtConsulta.Text;
                CreateCommand(query); // Chama o método CreateCommand da classe Connection
                using (SqlConnection connection = new SqlConnection(conn.connectionString))
                {

                    SqlCommand command = new SqlCommand(query, connection);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(dataTable);
                    MessageBox.Show("Consulta executada com sucesso!", "OK", MessageBoxButtons.OK);
                }
            }
            catch (SqlException erro)
            {
                MessageBox.Show("Erro ao se conectar no banco de dados \n" +
                "Verifique os dados informados" + erro);
            }
        }

        private void MontarTabelaComNomesEValores(DataTable dataTable)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (DataColumn column in dataTable.Columns)
                {
                    string nomeColuna = column.ColumnName;
                    string valorColuna = row[column].ToString();

                    tabelaComNomesEValores[nomeColuna] = valorColuna;
                }
            }
        }

        private void btnArquivo_Click(object sender, EventArgs e)
        {
            var fileStream = Metodos.BuscarArquivo();
            using (StreamReader reader = new StreamReader(fileStream))
            {
                conteudoArquivo = reader.ReadToEnd();
                txtArquivo.Text = Metodos.caminho;

            }
        }

        private void btnLista_Click(object sender, EventArgs e)
        {
            MontarTabelaComNomesEValores(dataTable);

            dataGridView2.Rows.Clear();

            dataGridView2.Columns.Add("Campo", "Campo");
            dataGridView2.Columns.Add("Valor", "Valor");

            foreach (var kvp in tabelaComNomesEValores)
            {
                dataGridView2.Rows.Add(kvp.Key, kvp.Value);
            }
        }
    }
}
