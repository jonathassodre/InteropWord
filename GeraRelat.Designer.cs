namespace InteropWord
{
    partial class GeraRelat
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
            btnArquivo = new Button();
            lblArquivo = new Label();
            txtArquivo = new TextBox();
            label1 = new Label();
            txtNovoArquivo = new TextBox();
            btnGerar = new Button();
            lblTermo = new Label();
            label2 = new Label();
            txtAntigo = new TextBox();
            txtNovo = new TextBox();
            menuStrip1 = new MenuStrip();
            btnSair = new Button();
            SuspendLayout();
            // 
            // btnArquivo
            // 
            btnArquivo.Location = new Point(424, 27);
            btnArquivo.Name = "btnArquivo";
            btnArquivo.Size = new Size(75, 23);
            btnArquivo.TabIndex = 0;
            btnArquivo.Text = "Procurar";
            btnArquivo.UseVisualStyleBackColor = true;
            btnArquivo.Click += btnArquivo_Click;
            // 
            // lblArquivo
            // 
            lblArquivo.AutoSize = true;
            lblArquivo.Location = new Point(12, 9);
            lblArquivo.Name = "lblArquivo";
            lblArquivo.Size = new Size(203, 15);
            lblArquivo.TabIndex = 1;
            lblArquivo.Text = "Selecione o Arquivo a ser modificado";
            // 
            // txtArquivo
            // 
            txtArquivo.Location = new Point(12, 27);
            txtArquivo.Name = "txtArquivo";
            txtArquivo.Size = new Size(406, 23);
            txtArquivo.TabIndex = 2;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 177);
            label1.Name = "label1";
            label1.Size = new Size(183, 15);
            label1.TabIndex = 3;
            label1.Text = "Informe o nome do novo arquivo";
            // 
            // txtNovoArquivo
            // 
            txtNovoArquivo.Location = new Point(12, 195);
            txtNovoArquivo.Name = "txtNovoArquivo";
            txtNovoArquivo.Size = new Size(183, 23);
            txtNovoArquivo.TabIndex = 4;
            // 
            // btnGerar
            // 
            btnGerar.Location = new Point(201, 195);
            btnGerar.Name = "btnGerar";
            btnGerar.Size = new Size(134, 23);
            btnGerar.TabIndex = 5;
            btnGerar.Text = "Gerar novo arquivo";
            btnGerar.UseVisualStyleBackColor = true;
            btnGerar.Click += btnGerar_Click;
            // 
            // lblTermo
            // 
            lblTermo.AutoSize = true;
            lblTermo.Location = new Point(12, 70);
            lblTermo.Name = "lblTermo";
            lblTermo.Size = new Size(183, 15);
            lblTermo.TabIndex = 6;
            lblTermo.Text = "Informe o termo a ser substituído";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 123);
            label2.Name = "label2";
            label2.Size = new Size(166, 15);
            label2.TabIndex = 7;
            label2.Text = "Informe o termo a ser inserido";
            // 
            // txtAntigo
            // 
            txtAntigo.Location = new Point(12, 88);
            txtAntigo.Name = "txtAntigo";
            txtAntigo.Size = new Size(183, 23);
            txtAntigo.TabIndex = 8;
            // 
            // txtNovo
            // 
            txtNovo.Location = new Point(12, 141);
            txtNovo.Name = "txtNovo";
            txtNovo.Size = new Size(183, 23);
            txtNovo.TabIndex = 9;
            // 
            // menuStrip1
            // 
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(506, 24);
            menuStrip1.TabIndex = 10;
            menuStrip1.Text = "menuStrip1";
            // 
            // btnSair
            // 
            btnSair.Location = new Point(424, 194);
            btnSair.Name = "btnSair";
            btnSair.Size = new Size(75, 23);
            btnSair.TabIndex = 11;
            btnSair.Text = "Sair";
            btnSair.UseVisualStyleBackColor = true;
            btnSair.Click += btnSair_Click;
            // 
            // GeraRelat
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(506, 233);
            Controls.Add(btnSair);
            Controls.Add(txtNovo);
            Controls.Add(txtAntigo);
            Controls.Add(label2);
            Controls.Add(lblTermo);
            Controls.Add(btnGerar);
            Controls.Add(txtNovoArquivo);
            Controls.Add(label1);
            Controls.Add(txtArquivo);
            Controls.Add(lblArquivo);
            Controls.Add(btnArquivo);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "GeraRelat";
            Text = "GeraRelat";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnArquivo;
        private Label lblArquivo;
        private TextBox txtArquivo;
        private Label label1;
        private TextBox txtNovoArquivo;
        private Button btnGerar;
        private Label lblTermo;
        private Label label2;
        private TextBox txtAntigo;
        private TextBox txtNovo;
        private MenuStrip menuStrip1;
        private Button btnSair;
    }
}