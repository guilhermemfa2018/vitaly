namespace Teste_Gerenciamento_de_Digitais_REP3
{
    partial class PopUp
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
            this.mensagem = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // mensagem
            // 
            this.mensagem.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.mensagem.BackColor = System.Drawing.Color.Transparent;
            this.mensagem.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mensagem.ForeColor = System.Drawing.Color.Red;
            this.mensagem.Location = new System.Drawing.Point(12, 314);
            this.mensagem.Name = "mensagem";
            this.mensagem.Size = new System.Drawing.Size(660, 37);
            this.mensagem.TabIndex = 0;
            this.mensagem.Text = "Texto de texte de mensagem";
            this.mensagem.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // PopUp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 362);
            this.Controls.Add(this.mensagem);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PopUp";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PopUp";
            this.Shown += new System.EventHandler(this.PopUp_Shown);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label mensagem;
    }
}