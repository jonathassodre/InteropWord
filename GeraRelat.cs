using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;

namespace InteropWord
{
    public partial class GeraRelat : Form
    {
        string conteudoArquivo = string.Empty;
        //string caminhoArquivo = string.Empty;
        //string nomeArquivo = string.Empty;
        Metodos metodos = new Metodos();

        public GeraRelat()
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

        private void btnGerar_Click(object sender, EventArgs e)
        {
            Metodos.CreateWordDocument(txtArquivo.Text, txtNovoArquivo.Text + ".docx", txtAntigo.Text, txtNovo.Text);
            MessageBox.Show("Processo Concluído", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSair_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
