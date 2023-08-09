using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InteropWord
{
    public partial class Start : Form
    {
        public Start()
        {
            InitializeComponent();
        }

        private void parâmetroÚnicoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeraRelat geraRelat = new GeraRelat();
            geraRelat.Show();
        }

        private void consultaSQLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GeraRelatSql geraRelatSql = new GeraRelatSql();
            geraRelatSql.Show();
        }

        private void sairToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
