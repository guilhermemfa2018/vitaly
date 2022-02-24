using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace Teste_Gerenciamento_de_Digitais_REP3
{
    public partial class PopUp : Form
    {
        int timeOut=1;
        public PopUp(string texto, Image fundo, int timeOut)
        {
            InitializeComponent();
            this.Width = fundo.Width;
            this.Height = fundo.Height;
            this.BackgroundImageLayout = ImageLayout.Center;
            this.BackgroundImage = fundo;
            this.timeOut = timeOut;
            mensagem.Text = texto;
        }

        private void PopUp_Shown(object sender, EventArgs e)
        {
            this.Refresh();
            Thread.Sleep(timeOut);
            this.Close();
        }

    }
}
