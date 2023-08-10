using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;
using System.Web.UI;
using Microsoft.Office.Interop;
using static InteropWord.Connection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace InteropWord
{
    public partial class GeraRelatSql : Form
    {

        Connection conn = new Connection();
        string conteudoArquivo = string.Empty;
        string caminhoArquivo = string.Empty;
        string nomeArquivo = string.Empty;
        private Dictionary<string, string> tabelaComNomesEValores = new Dictionary<string, string>();
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        List<object> rowData = new List<object>();


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
            try
            {
                query = txtConsulta.Text;
                CreateCommand(query); // Chama o método CreateCommand da classe Connection
                using (SqlConnection connection = new SqlConnection(conn.connectionString))
                {

                    SqlCommand command = new SqlCommand(query, connection);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(dataTable);
                    //MessageBox.Show("Consulta executada com sucesso!", "OK", MessageBoxButtons.OK);
                }
            }
            catch (SqlException erro)
            {
                MessageBox.Show("Erro ao se conectar no banco de dados \n" +
                "Verifique os dados informados" + erro);
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
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            MontarTabelaComNomesEValores(dataTable);

            // Adicione colunas ao DataGridView com os nomes dos campos da consulta
            foreach (DataColumn column in dataTable.Columns)
            {
                dataGridView2.Columns.Add(column.ColumnName, column.ColumnName);
            }

            foreach (DataRow row in dataTable.Rows)
            {
                dataGridView2.Rows.Add(row.ItemArray);
            }
        }

        private void btnGerar_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow linha in dataGridView2.Rows)
            {
                for (int columnIndex = 0; columnIndex < dataGridView2.Columns.Count; columnIndex++)
                {
                    string header = dataGridView2.Columns[columnIndex].HeaderText;
                    string cellValue = linha.Cells[columnIndex].Value.ToString();

                    FindAndReplace(wordApp, txtArquivo.Text, header, cellValue);

                }
                MessageBox.Show($"Processo concluído para o colaborador {dataTable.Rows[0]}", "Sucesso", MessageBoxButtons.OK);

                string novoArquivo = $"{txtNovoArquivo.Text}_{linha.Index}.docx";
                Metodos.CreateWordDocument(txtArquivo.Text, novoArquivo);
            }

            MessageBox.Show("Processo concluído!", "Sucesso", MessageBoxButtons.OK);
        }

        public static void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object filename, object header, object linha)
        {
            try
            {
                object missing = Missing.Value;

                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;

                    wordApp.Visible = false;
                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();
                    wordApp.Selection.Find.Execute(ref header, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing,
                                                   ref missing, ref missing, ref linha,
                                                   ref missing, ref missing,
                                                   ref missing, ref missing,
                                                   ref missing);

                    myWordDoc.Close();
                    wordApp.Quit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }
        }
    }
}
