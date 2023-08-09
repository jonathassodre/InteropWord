namespace InteropWord
{
    partial class Start
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
            menuStrip1 = new MenuStrip();
            geradoresToolStripMenuItem = new ToolStripMenuItem();
            parâmetroÚnicoToolStripMenuItem = new ToolStripMenuItem();
            consultaSQLToolStripMenuItem = new ToolStripMenuItem();
            sairToolStripMenuItem = new ToolStripMenuItem();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { geradoresToolStripMenuItem, sairToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(311, 24);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // geradoresToolStripMenuItem
            // 
            geradoresToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { parâmetroÚnicoToolStripMenuItem, consultaSQLToolStripMenuItem });
            geradoresToolStripMenuItem.Name = "geradoresToolStripMenuItem";
            geradoresToolStripMenuItem.Size = new Size(72, 20);
            geradoresToolStripMenuItem.Text = "Geradores";
            // 
            // parâmetroÚnicoToolStripMenuItem
            // 
            parâmetroÚnicoToolStripMenuItem.Name = "parâmetroÚnicoToolStripMenuItem";
            parâmetroÚnicoToolStripMenuItem.Size = new Size(180, 22);
            parâmetroÚnicoToolStripMenuItem.Text = "Parâmetro Único";
            parâmetroÚnicoToolStripMenuItem.Click += parâmetroÚnicoToolStripMenuItem_Click;
            // 
            // consultaSQLToolStripMenuItem
            // 
            consultaSQLToolStripMenuItem.Name = "consultaSQLToolStripMenuItem";
            consultaSQLToolStripMenuItem.Size = new Size(180, 22);
            consultaSQLToolStripMenuItem.Text = "Consulta SQL";
            consultaSQLToolStripMenuItem.Click += consultaSQLToolStripMenuItem_Click;
            // 
            // sairToolStripMenuItem
            // 
            sairToolStripMenuItem.Name = "sairToolStripMenuItem";
            sairToolStripMenuItem.Size = new Size(38, 20);
            sairToolStripMenuItem.Text = "Sair";
            sairToolStripMenuItem.Click += sairToolStripMenuItem_Click;
            // 
            // Start
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(311, 46);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "Start";
            Text = "Gerador de Relatórios via Word";
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem geradoresToolStripMenuItem;
        private ToolStripMenuItem parâmetroÚnicoToolStripMenuItem;
        private ToolStripMenuItem consultaSQLToolStripMenuItem;
        private ToolStripMenuItem sairToolStripMenuItem;
    }
}