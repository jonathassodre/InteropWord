namespace InteropWord
{
    partial class GeraRelatSql
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GeraRelatSql));
            lblConsulta = new Label();
            txtConsulta = new TextBox();
            connectionBindingSource = new BindingSource(components);
            txtArquivo = new TextBox();
            lblArquivo = new Label();
            btnArquivo = new Button();
            btnLista = new Button();
            dataGridView2 = new DataGridView();
            btnGerar = new Button();
            btnLimpa = new Button();
            ((System.ComponentModel.ISupportInitialize)connectionBindingSource).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            SuspendLayout();
            // 
            // lblConsulta
            // 
            lblConsulta.AutoSize = true;
            lblConsulta.Location = new Point(12, 53);
            lblConsulta.Name = "lblConsulta";
            lblConsulta.Size = new Size(78, 15);
            lblConsulta.TabIndex = 1;
            lblConsulta.Text = "Consulta SQL";
            // 
            // txtConsulta
            // 
            txtConsulta.Location = new Point(121, 71);
            txtConsulta.Name = "txtConsulta";
            txtConsulta.ScrollBars = ScrollBars.Vertical;
            txtConsulta.Size = new Size(752, 23);
            txtConsulta.TabIndex = 2;
            // 
            // connectionBindingSource
            // 
            connectionBindingSource.DataSource = typeof(Connection);
            // 
            // txtArquivo
            // 
            txtArquivo.Location = new Point(12, 27);
            txtArquivo.Name = "txtArquivo";
            txtArquivo.Size = new Size(442, 23);
            txtArquivo.TabIndex = 6;
            // 
            // lblArquivo
            // 
            lblArquivo.AutoSize = true;
            lblArquivo.Location = new Point(12, 9);
            lblArquivo.Name = "lblArquivo";
            lblArquivo.Size = new Size(203, 15);
            lblArquivo.TabIndex = 5;
            lblArquivo.Text = "Selecione o Arquivo a ser modificado";
            // 
            // btnArquivo
            // 
            btnArquivo.Location = new Point(460, 27);
            btnArquivo.Name = "btnArquivo";
            btnArquivo.Size = new Size(75, 23);
            btnArquivo.TabIndex = 4;
            btnArquivo.Text = "Procurar";
            btnArquivo.UseVisualStyleBackColor = true;
            btnArquivo.Click += btnArquivo_Click;
            // 
            // btnLista
            // 
            btnLista.Location = new Point(879, 71);
            btnLista.Name = "btnLista";
            btnLista.Size = new Size(148, 23);
            btnLista.TabIndex = 7;
            btnLista.Text = "Lista Campos e Valores";
            btnLista.UseVisualStyleBackColor = true;
            btnLista.Click += btnLista_Click;
            // 
            // dataGridView2
            // 
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Location = new Point(12, 100);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.RowTemplate.Height = 25;
            dataGridView2.Size = new Size(1015, 331);
            dataGridView2.TabIndex = 8;
            // 
            // btnGerar
            // 
            btnGerar.Location = new Point(893, 437);
            btnGerar.Name = "btnGerar";
            btnGerar.Size = new Size(134, 23);
            btnGerar.TabIndex = 11;
            btnGerar.Text = "Gerar novo arquivo";
            btnGerar.UseVisualStyleBackColor = true;
            btnGerar.Click += btnGerar_Click;
            // 
            // btnLimpa
            // 
            btnLimpa.Location = new Point(12, 71);
            btnLimpa.Name = "btnLimpa";
            btnLimpa.Size = new Size(103, 23);
            btnLimpa.TabIndex = 12;
            btnLimpa.Text = "Limpa Consulta";
            btnLimpa.UseVisualStyleBackColor = true;
            btnLimpa.Click += btnLimpa_Click;
            // 
            // GeraRelatSql
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1049, 471);
            Controls.Add(btnLimpa);
            Controls.Add(btnGerar);
            Controls.Add(dataGridView2);
            Controls.Add(btnLista);
            Controls.Add(txtArquivo);
            Controls.Add(lblArquivo);
            Controls.Add(btnArquivo);
            Controls.Add(txtConsulta);
            Controls.Add(lblConsulta);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "GeraRelatSql";
            Text = "Gerador de Relatórios via Word";
            ((System.ComponentModel.ISupportInitialize)connectionBindingSource).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Label lblConsulta;
        private TextBox txtConsulta;
        private BindingSource connectionBindingSource;
        private TextBox txtArquivo;
        private Label lblArquivo;
        private Button btnArquivo;
        private Button btnLista;
        private DataGridView dataGridView2;
        private Button btnGerar;
        private Button btnLimpa;
    }
}