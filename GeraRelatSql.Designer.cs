﻿namespace InteropWord
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
            btnTestar = new Button();
            lblConsulta = new Label();
            txtConsulta = new TextBox();
            connectionBindingSource = new BindingSource(components);
            txtArquivo = new TextBox();
            lblArquivo = new Label();
            btnArquivo = new Button();
            btnLista = new Button();
            dataGridView2 = new DataGridView();
            btnGerar = new Button();
            txtNovoArquivo = new TextBox();
            label1 = new Label();
            ((System.ComponentModel.ISupportInitialize)connectionBindingSource).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            SuspendLayout();
            // 
            // btnTestar
            // 
            btnTestar.Location = new Point(460, 71);
            btnTestar.Name = "btnTestar";
            btnTestar.Size = new Size(75, 23);
            btnTestar.TabIndex = 0;
            btnTestar.Text = "Executa";
            btnTestar.UseVisualStyleBackColor = true;
            btnTestar.Click += btnTestar_Click;
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
            txtConsulta.Location = new Point(12, 71);
            txtConsulta.Name = "txtConsulta";
            txtConsulta.ScrollBars = ScrollBars.Vertical;
            txtConsulta.Size = new Size(442, 23);
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
            btnLista.Location = new Point(12, 100);
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
            dataGridView2.Location = new Point(12, 129);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.RowTemplate.Height = 25;
            dataGridView2.Size = new Size(523, 302);
            dataGridView2.TabIndex = 8;
            // 
            // btnGerar
            // 
            btnGerar.Location = new Point(201, 452);
            btnGerar.Name = "btnGerar";
            btnGerar.Size = new Size(134, 23);
            btnGerar.TabIndex = 11;
            btnGerar.Text = "Gerar novo arquivo";
            btnGerar.UseVisualStyleBackColor = true;
            btnGerar.Click += btnGerar_Click;
            // 
            // txtNovoArquivo
            // 
            txtNovoArquivo.Location = new Point(12, 452);
            txtNovoArquivo.Name = "txtNovoArquivo";
            txtNovoArquivo.Size = new Size(183, 23);
            txtNovoArquivo.TabIndex = 10;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 434);
            label1.Name = "label1";
            label1.Size = new Size(183, 15);
            label1.TabIndex = 9;
            label1.Text = "Informe o nome do novo arquivo";
            // 
            // GeraRelatSql
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(547, 491);
            Controls.Add(btnGerar);
            Controls.Add(txtNovoArquivo);
            Controls.Add(label1);
            Controls.Add(dataGridView2);
            Controls.Add(btnLista);
            Controls.Add(txtArquivo);
            Controls.Add(lblArquivo);
            Controls.Add(btnArquivo);
            Controls.Add(txtConsulta);
            Controls.Add(lblConsulta);
            Controls.Add(btnTestar);
            Name = "GeraRelatSql";
            Text = "GeraRelatSql";
            ((System.ComponentModel.ISupportInitialize)connectionBindingSource).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnTestar;
        private Label lblConsulta;
        private TextBox txtConsulta;
        private BindingSource connectionBindingSource;
        private TextBox txtArquivo;
        private Label lblArquivo;
        private Button btnArquivo;
        private Button btnLista;
        private DataGridView dataGridView2;
        private Button btnGerar;
        private TextBox txtNovoArquivo;
        private Label label1;
    }
}