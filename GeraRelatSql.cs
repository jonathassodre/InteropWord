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
        string query = "";
        DataTable dataTable = new DataTable();

        public GeraRelatSql()
        {
            InitializeComponent();
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

        //private void btnTestar_Click(object sender, EventArgs e)
        // {
        //     try
        //     {
        //         query = txtConsulta.Text;
        //         CreateCommand(query); // Chama o método CreateCommand da classe Connection
        //         using (SqlConnection connection = new SqlConnection(conn.connectionString))
        //         {

        //             SqlCommand command = new SqlCommand(query, connection);
        //             SqlDataAdapter adapter = new SqlDataAdapter(command);
        //             adapter.Fill(dataTable);
        //             MessageBox.Show("Consulta executada com sucesso!", "OK", MessageBoxButtons.OK);
        //         }
        //     }
        //     catch (SqlException erro)
        //     {
        //         MessageBox.Show("Erro ao se conectar no banco de dados \n" +
        //         "Verifique os dados informados" + erro);
        //     }
        // }

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
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                // Obtenha o cabeçalho da coluna do DataGridView
                List<string> cabecalho = new();
                List<string> linha = new();

                foreach (DataGridViewColumn coluna in dataGridView2.Columns)
                {
                    cabecalho.Add(coluna.HeaderText);
                    linha.Add(row.Cells[coluna.Index].Value.ToString());


                }
                string novoNomeArquivo = Path.Combine(Path.GetDirectoryName(txtArquivo.Text), $"Contrato {linha.First()}.docx");

                CreateWordDocument(txtArquivo.Text, novoNomeArquivo, cabecalho, linha);

            }

            MessageBox.Show("Processo concluído!", "Sucesso", MessageBoxButtons.OK);
        }



        public static void CreateWordDocument(object filename, object SaveAs, List<string> cabecalho, List<string> linha)
        {
            try
            {
                object textoCabecalho = "";
                object textoLinha = "";

                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = Missing.Value;

                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;

                    object isvisible = false;

                    wordApp.Visible = false;
                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();

                    for (int i = 0; i < cabecalho.Count; i++)
                    {
                        textoCabecalho = cabecalho[i];
                        textoLinha = linha[i];

                        FindAndReplace(wordApp, textoCabecalho, textoLinha);
                    }



                    myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing);

                    myWordDoc.Close();
                    wordApp.Quit();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }

        }
        public static void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object cabecalho, object linha)
        {
            try
            {
                object matchCase = false;
                object matchwholeWord = false;
                object matchwildCards = false;
                object matchSoundLike = false;
                object nmatchAllforms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiactitics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = -2;
                object wrap = 1;

                wordApp.Selection.Find.Execute(ref cabecalho, ref matchCase,
                                                ref matchwholeWord, ref matchwildCards, ref matchSoundLike,
                                                ref nmatchAllforms, ref forward,
                                                ref wrap, ref format, ref linha,
                                                    ref replace, ref matchKashida,
                                                ref matchDiactitics, ref matchAlefHamza,
                                                 ref matchControl);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }
        }
    }
}
