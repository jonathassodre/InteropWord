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
using System.Globalization;
using Microsoft.AspNetCore.Http;

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

        private void MontarTabelaComNomesEValores(DataTable dataTable)
        {
            try
            {
                dataTable.Clear();

                query = txtConsulta.Text;
                CreateCommand(query); // Chama o método CreateCommand da classe Connection
                using (SqlConnection connection = new SqlConnection(conn.connectionString))
                {
                    try {
                        SqlCommand command = new SqlCommand(query, connection);
                        SqlDataAdapter adapter = new SqlDataAdapter(command);
                        adapter.Fill(dataTable);
                    }
                    catch (System.InvalidOperationException erro) {
                        MessageBox.Show("Informe uma consulta SQL!", "Erro", MessageBoxButtons.OKCancel);
                    }  
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
            if(txtArquivo.Text == "")
            {
                MessageBox.Show("Nenhum arquivo foi selecionado.", "Erro", MessageBoxButtons.OKCancel);
                return;
            }
            if(dataGridView2.Rows.Count == 0 || dataGridView2.Rows[0] == null)
            {
                MessageBox.Show("Nenhuma consulta SQL foi executada.", "Erro", MessageBoxButtons.OKCancel);
                return;
            }

            GeraArquivo(dataGridView2, txtArquivo.Text);
        }

        public static void GeraArquivo(DataGridView dataGridView, string arquivoOrigem)
        {
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // Obtenha o cabeçalho da coluna do DataGridView
                List<string> cabecalho = new();
                List<string> linha = new();

                bool linhaVazia = true;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    {
                        linhaVazia = false;
                        break; // Não é uma linha vazia, pode continuar processando
                    }
                }

                foreach (DataGridViewColumn coluna in dataGridView.Columns)
                {
                    if (linhaVazia)
                    {
                        // Interrompa a geração, pois a linha está vazia
                        MessageBox.Show("Processo concluído!", "Sucesso", MessageBoxButtons.OK);
                        return;
                    }

                    cabecalho.Add(coluna.HeaderText);
                    linha.Add(row.Cells[coluna.Index].Value.ToString());
                }
                string novoNomeArquivo = Path.Combine(Path.GetDirectoryName(arquivoOrigem), $"Contrato {linha.First()}.docx");

                CreateWordDocument(arquivoOrigem, novoNomeArquivo, cabecalho, linha);
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

                switch (cabecalho)
                {
                    case "FILENDCEP":
                        linha = $"{linha.ToString().Substring(0, 5)}-{linha.ToString().Substring(5)}";
                        break;
                    case "FILCGCCOMPLETO":
                        linha = $"{linha.ToString().Substring(0, 2)}.{linha.ToString().Substring(2, 3)}.{linha.ToString().Substring(5, 3)}/{linha.ToString().Substring(8, 4)}-{linha.ToString().Substring(12)}";
                        break;
                    case "PFFVALORSALARIO":
                        decimal salario = decimal.Parse(linha.ToString()); // Supondo que a linha seja uma string que representa o valor
                        linha = salario.ToString("C");
                        break;
                    case "HORAINICIOTURNO1":
                        linha = $"{linha.ToString().Substring(0, 2)}:{linha.ToString().Substring(2)}";
                        break;
                    case "HORAFIMTURNO2":
                        linha = $"{linha.ToString().Substring(0, 2)}:{linha.ToString().Substring(2)}";
                        break;
                    case "PFUDTINICIOCONTRATO":
                        DateTime data = DateTime.ParseExact(linha.ToString(), "dd/MM/yyyy hh:mm:ss", new CultureInfo("pt-BR"));
                        string dataExtenso = data.ToString("dd 'de' MMMM 'de' yyyy", new CultureInfo("pt-BR"));
                        linha = dataExtenso;
                        break;
                    case "FILENDCIDADE":
                        linha = linha.ToString();
                        break;
                }

                //wordApp.Selection.Find.Execute(ref cabecalho, ref matchCase,
                //                                ref matchwholeWord, ref matchwildCards, ref matchSoundLike,
                //                                ref nmatchAllforms, ref forward,
                //                                ref wrap, ref format, ref linha,
                //                                    ref replace, ref matchKashida,
                //                                ref matchDiactitics, ref matchAlefHamza,
                //                                 ref matchControl);

                bool encontrou = true;

                while (encontrou)
                {
                    encontrou = wordApp.Selection.Find.Execute(
                        ref cabecalho, ref matchCase,
                        ref matchwholeWord, ref matchwildCards, ref matchSoundLike,
                        ref nmatchAllforms, ref forward,
                        ref wrap, ref format, ref linha,
                        ref replace, ref matchKashida,
                        ref matchDiactitics, ref matchAlefHamza,
                        ref matchControl);

                    if (encontrou)
                    {
                        // Substitua o valor na seleção
                        wordApp.Selection.Text = linha.ToString();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Aviso!", MessageBoxButtons.OK);
            }
        }

        private void btnLimpa_Click(object sender, EventArgs e)
        {
            txtConsulta.Text = string.Empty;
        }
    }
}
