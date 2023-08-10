using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
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
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

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

        private void btnGerar_Click(object sender, EventArgs e)
        {

            FindAndReplace(wordApp, tabelaComNomesEValores, txtArquivo.Text);

            Metodos.CreateWordDocument(txtArquivo.Text, txtNovoArquivo.Text + ".docx");

            MessageBox.Show("Processo Finalizado! Arquivo: " + txtNovoArquivo.Text + " salvo com sucesso!", "Aviso", MessageBoxButtons.OK);
        }

        public static void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, Dictionary<string, string> tabelaComNomesEValores, object filename)
        {
            try
            {
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
                    object toFindText;
                    object replaceWithText;

                    foreach (var kvp in tabelaComNomesEValores)
                    {
                        toFindText = kvp.Key;
                        replaceWithText = kvp.Value;

                        wordApp.Selection.Find.Execute(ref toFindText, ref matchCase,
                                                    ref matchwholeWord, ref matchwildCards, ref matchSoundLike,

                                                    ref nmatchAllforms, ref forward,

                                                    ref wrap, ref format, ref replaceWithText,

                                                        ref replace, ref matchKashida,

                                                    ref matchDiactitics, ref matchAlefHamza,

                                                     ref matchControl);
                    }

                    //myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                    //                                                ref missing, ref missing, ref missing,
                    //                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    //                                                ref missing, ref missing, ref missing);

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
