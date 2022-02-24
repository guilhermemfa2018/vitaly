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
    public partial class GerenciarDigitaisREP3 : Form
    {
        CKREP3.Comunicador comunicador = CKREP3.Comunicador.Instance;
        public GerenciarDigitaisREP3()
        {
            InitializeComponent();
            //UCBioBSP.MetodosExportados.gUtilizarVirdi = false;
        }
        bool ocupado = false;
        string copia = "";


        private void btLerTempIndividual_Click(object sender, EventArgs e)
        {
            if (ocupado)
            {
                MessageBox.Show("Software ocupado");
                return;
            }
            ocupado = true;
            try
            {
                cxLerTemplate1.Text = "";
                cxLerTemplate2.Text = "";
                cxLerTemplate3.Text = "";
                cxLerTemplate4.Text = "";
                cxLerTemplate5.Text = "";

                if (cxNumPIS.Text.Length == 0)
                {
                    MessageBox.Show("Informe o numero do crachá");
                    return;
                }
                string[] templates;
                int ret = comunicador.LerTemplatesUsuario(cxNumPIS.Text, out templates);

                if (templates != null)
                {
                    int quantTemplates = templates.Length;
                    if (quantTemplates > 0) cxLerTemplate1.Text = templates[0];
                    if (quantTemplates > 1) cxLerTemplate2.Text = templates[1];
                    if (quantTemplates > 2) cxLerTemplate3.Text = templates[2];
                    if (quantTemplates > 3) cxLerTemplate4.Text = templates[3];
                    if (quantTemplates > 4) cxLerTemplate5.Text = templates[4];
                }
                if (ret != CKREP3.Defs.RetornoDefs.RETORNO_OK)
                {
                    string msg = "";
                    MessageBox.Show(msg + "Ocorreu o Erro Cod.: " + ret);
                    return;
                }
            }
            finally
            {
                ocupado = false;
            }
        }


        private void btGravarTempIndividual_Click(object sender, EventArgs e)
        {
            if(ocupado) { 
                MessageBox.Show("Software ocupado"); 
                return;
            } 
            ocupado = true;
            try
            {
                if (cxNumPIS.Text.Length == 0)
                {
                    MessageBox.Show("Informe o numero do crachá");
                    return;
                }
                if (cxGravarTemplate1.Text.Length == 0)
                {
                    MessageBox.Show("Primeiro template não preenchido");
                    return;
                }
            
                List<string> TemplatesTemp = new List<string>();
                if (cxGravarTemplate1.Text.Length > 0) TemplatesTemp.Add(cxGravarTemplate1.Text);
                if (cxGravarTemplate2.Text.Length > 0) TemplatesTemp.Add(cxGravarTemplate2.Text);
                if (cxGravarTemplate3.Text.Length > 0) TemplatesTemp.Add(cxGravarTemplate3.Text);
                if (cxGravarTemplate4.Text.Length > 0) TemplatesTemp.Add(cxGravarTemplate4.Text);
                if (cxGravarTemplate5.Text.Length > 0) TemplatesTemp.Add(cxGravarTemplate5.Text);

                string[] linhasTemplates = TemplatesTemp.ToArray();
                int ret = comunicador.GravarTemplatesUsuario(cxNumPIS.Text, linhasTemplates);
                if (ret != CKREP3.Defs.RetornoDefs.RETORNO_OK)
                {
                    string msg = "";
                    MessageBox.Show(msg + "Ocorreu o Erro Cod.: " + ret);
                    return;
                }
            }
            finally
            {
                ocupado = false;
            }
        }

        private void btGravarTempConcatenado_Click(object sender, EventArgs e)
        {
            if(ocupado) 
            { 
                MessageBox.Show("Software ocupado"); 
                return;
            } 
            ocupado = true;
            try
            {
                if(cxNumPIS.Text.Length == 0)
                {
                    MessageBox.Show("Informe o numero do crachá");
                    return;
                }
                if(cxGravarTemplate1.Text.Length == 0)
                {
                    MessageBox.Show("Primeiro template não preenchido");
                    return;
                }
                int ret = 0;//ALEFequipamento.GravarTemplatesUsuario(cxNumCracha.Text, cxGravarTemplate1.Text);
                //ALEF if (ret != RetornoFuncoes.RETORNO_OK)
                {
                    string msg = "";
                    //ALEF if (ret == RetornoFuncoes.ERRO_MODO_GB_NAO_ATIVADO) msg = "Modo de degerenciamento desativado. ";
                    MessageBox.Show(msg +"Ocorreu o Erro Cod.: " + ret);
                    return;
                }
            }
            finally
            {
                ocupado = false;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if(ocupado) 
            { 
                MessageBox.Show("Software ocupado"); 
                return;
            } 
            ocupado = true;
            try
            {
                int ret = 0;//ALEFequipamento.AtivarGerencimentoBiometrico();
                //ALEF if (ret != RetornoFuncoes.RETORNO_OK)
                {
                    MessageBox.Show("Ocorreu o Erro Cod.: " + ret);
                    return;
                }
            }
            finally
            {
                ocupado = false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(ocupado) 
            { 
                MessageBox.Show("Software ocupado"); 
                return;
            } 
            ocupado = true;
            try
            {
                int ret = 0;//ALEFequipamento.DesativarGerencimentoBiometrico();
                //ALEF if (ret != RetornoFuncoes.RETORNO_OK)
                {
                    MessageBox.Show("Ocorreu o Erro Cod.: " + ret);
                    return;
                }
            }
            finally
            {
                ocupado = false;
            }
        }
        private int listarUsuariosREP()
        {
            TBListaUsuarios.Rows.Clear();
            KeyValuePair<string, string>[] lista;
            int ret = comunicador.ListarUsuariosBiometria(out  lista);
            TBListaUsuarios.Rows.Clear();
            if (ret != CKREP3.Defs.RetornoDefs.RETORNO_OK)
            {
                MessageBox.Show("Ocorreu o Erro Cod.: " + ret);
                return ret;
            }
            foreach (KeyValuePair<string, string> linha in lista)
            {
                TBListaUsuarios.Rows.Add(linha.Key, linha.Value);
            }
            if (TBListaUsuarios.RowCount > 0) TBListaUsuarios.Rows[0].Selected = true;

            if (TBListaUsuarios.SelectedRows.Count == 1)
            {
                string cracha = TBListaUsuarios.SelectedRows[0].Cells[0].Value.ToString();
                cxNumPIS.Text = cracha;
            }
            return CKREP3.Defs.RetornoDefs.RETORNO_OK; 
        }

        private void btListarUsuarios_Click(object sender, EventArgs e)
        {
            if(ocupado) 
            { 
                MessageBox.Show("Software ocupado");
                return;
            }
            ocupado = true;
            try
            {
               listarUsuariosREP();
            }
            finally
            {
                ocupado = false;
            }
        }

        private void btExcluirUsuario_Click(object sender, EventArgs e)
        {
            if(ocupado) 
            { 
                MessageBox.Show("Software ocupado"); 
                return;
            } 
            ocupado = true;
            try
            {
                if (TBListaUsuarios.SelectedRows.Count == 1)
                {
                    string PISASerRemovido = TBListaUsuarios.SelectedRows[0].Cells[0].Value.ToString();
                    int ret = 0;
                    comunicador.ExcluirTemplatesUsuario(PISASerRemovido);
                    if (ret != CKREP3.Defs.RetornoDefs.RETORNO_OK)
                    {
                        MessageBox.Show("Ocorreu o Erro Cod.: " + ret);
                        return;
                    }
                    MessageBox.Show("Excluido com sucesso!", "Sucesso");
                }
            }
            finally
            {
                ocupado = false;
            }
        }

        private void TBListaUsuarios_Click(object sender, EventArgs e)
        {
            if (TBListaUsuarios.SelectedRows.Count == 1)
            {
                string cracha = TBListaUsuarios.SelectedRows[0].Cells[0].Value.ToString();
                cxNumPIS.Text = cracha;
            }
        }

        private void GerenciarDigitaisTupa2_Load(object sender, EventArgs e)
        {

        }

        private void GerenciarDigitaisTupa2_FormClosed(object sender, FormClosedEventArgs e)
        {
            //ALEFequipamento.DesativarGerencimentoBiometrico();
        }

        private void btGravarTempConcatenado_MouseHover(object sender, EventArgs e)
        {
            cxGravarTemplate2.Visible = false;
            cxGravarTemplate3.Visible = false;
            cxGravarTemplate4.Visible = false;
            cxGravarTemplate5.Visible = false;
        }

        private void btGravarTempConcatenado_MouseLeave(object sender, EventArgs e)
        {
            cxGravarTemplate2.Visible = true;
            cxGravarTemplate3.Visible = true;
            cxGravarTemplate4.Visible = true;
            cxGravarTemplate5.Visible = true;
        }

        private void btcp_cxLerTemplate1_Click(object sender, EventArgs e)
        {
            Button bt = (Button) sender;
            string[] nome = bt.Name.Split('_');
            TextBox campo = this.Controls.Find(nome[1], true).FirstOrDefault() as TextBox;
            if (campo.Text.Length>0)
            {
                copia = campo.Text;
                TxCopiado.Visible = true;
                this.Refresh();
                Thread.Sleep(350);
            }
            TxCopiado.Visible = false;
            this.Refresh();
        }

        private void CpCL_cxGravarTemplate1_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            string[] nome = bt.Name.Split('_');
            TextBox campo = this.Controls.Find(nome[1], true).FirstOrDefault() as TextBox;
            if(copia.Length>0) campo.Text = copia;
            copia = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cxGravarTemplate1.Text = "";
            cxGravarTemplate2.Text = "";
            cxGravarTemplate3.Text = "";
            cxGravarTemplate4.Text = "";
            cxGravarTemplate5.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cxLerTemplate1.Text = "";
            cxLerTemplate2.Text = "";
            cxLerTemplate3.Text = "";
            cxLerTemplate4.Text = "";
            cxLerTemplate5.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {

            string template;
            int erro = comunicador.Capturar1Digital(pictureBox1.Handle, 120, out template);
            if (erro != CKREP3.Defs.RetornoDefs.RETORNO_OK)
            {
                if (erro == CKREP3.Defs.RetornoDefs.ERRO_LEITOR_BIO_USB_NAO_ENCONTRADO)
                {
                    PopUp p = new PopUp("Por favor, conecte o leitor USB", Teste_Gerenciamento_de_Digitais_REP3.Properties.Resources.ConecteLeitorUsb, 4000);
                    p.ShowDialog();
                    return;
                }
                MessageBox.Show("Erro:" + erro.ToString());
            }
            else
            {
                if (template != null)
                {
                    QuantTemplatesLidos.Text = "1";
                    cxCadTemplate1.Text = template;
                    cxCadTemplate2.Text = "";
                    cxCadTemplate3.Text = "";
                    cxCadTemplate4.Text = "";
                    cxCadTemplate5.Text = "";
                    QuantTemplatesLidos.Text = "";
                    cxCadNTemplate1.Text = "";
                    cxCadNTemplate2.Text = "";
                    cxCadNTemplate3.Text = "";
                    cxCadNTemplate4.Text = "";
                    cxCadNTemplate5.Text = "";
                }
            }
            pictureBox1.Image = null;
        }
        
        private void button4_Click(object sender, EventArgs e)
        {
            string[] templates;
            int erro = comunicador.CapturarDigitais(out templates);
            if (erro != CKREP3.Defs.RetornoDefs.RETORNO_OK)
            {
                if (erro == CKREP3.Defs.RetornoDefs.ERRO_LEITOR_BIO_USB_NAO_ENCONTRADO)
                {
                    PopUp p = new PopUp("Por favor, conecte o leitor USB", Teste_Gerenciamento_de_Digitais_REP3.Properties.Resources.ConecteLeitorUsb, 4000);
                    p.ShowDialog();
                    return;
                }
                MessageBox.Show("Erro:" + erro.ToString());
            }
            else
            {
                if (templates != null)
                {
                    cxCadTemplate1.Text = "";
                    cxCadTemplate2.Text = "";
                    cxCadTemplate3.Text = "";
                    cxCadTemplate4.Text = "";
                    cxCadTemplate5.Text = "";
                    QuantTemplatesLidos.Text = "";
                    cxCadNTemplate1.Text = "";
                    cxCadNTemplate2.Text = "";
                    cxCadNTemplate3.Text = "";
                    cxCadNTemplate4.Text = "";
                    cxCadNTemplate5.Text = "";

                    int quantTemplates = templates.Length;
                    QuantTemplatesLidos.Text = quantTemplates.ToString();
                    if (quantTemplates > 0) cxCadTemplate1.Text = templates[0];
                    if (quantTemplates > 1) cxCadTemplate2.Text = templates[1];
                    if (quantTemplates > 2) cxCadTemplate3.Text = templates[2];
                    if (quantTemplates > 3) cxCadTemplate4.Text = templates[3];
                    if (quantTemplates > 4) cxCadTemplate5.Text = templates[4];

                    //Exemplo de como pegar de qual dedo foi capturado do leitor
                    //No Rep não é possivel saber o dedo, então a numeração é dada na ordem do cadastro
                    if (quantTemplates > 0) cxCadNTemplate1.Text = Convert.FromBase64String(templates[0])[0].ToString();
                    if (quantTemplates > 1) cxCadNTemplate2.Text = Convert.FromBase64String(templates[1])[0].ToString();
                    if (quantTemplates > 2) cxCadNTemplate3.Text = Convert.FromBase64String(templates[2])[0].ToString();
                    if (quantTemplates > 3) cxCadNTemplate4.Text = Convert.FromBase64String(templates[3])[0].ToString();
                    if (quantTemplates > 4) cxCadNTemplate5.Text = Convert.FromBase64String(templates[4])[0].ToString();
                }
            }
        }

        private void verificarConexao_Click(object sender, EventArgs e)
        {
            

            int retorno = comunicador.InicializarGerenciamentoBio(cxNumSerie.Text, cxSenha.Text, cxCPF.Text, cxIp.Text, cxPorta.Text); 
            if (retorno != CKREP3.Defs.RetornoDefs.RETORNO_OK)
            {
                MessageBox.Show("Falha ao iniciar CKREP3, erro cod.: " + retorno.ToString("X4"), "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            retorno = listarUsuariosREP();
            if (retorno != CKREP3.Defs.RetornoDefs.RETORNO_OK)
            {
                MessageBox.Show("Falha ao listar usuarios CKREP3, erro cod.: " + retorno.ToString("X4"), "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btLerTempIndividual.Enabled = true;
            btGravarTempIndividual.Enabled = true;
            btGravarTempConcatenado.Enabled = true;
            btExcluirUsuario.Enabled = true;
            btListarUsuarios.Enabled = true;
            btExportarFunc.Enabled = true;

            btcp_cxLerTemplate1.Enabled = true;
            btcp_cxLerTemplate2.Enabled = true;
            btcp_cxLerTemplate3.Enabled = true;
            btcp_cxLerTemplate4.Enabled = true;
            btcp_cxLerTemplate5.Enabled = true;

            CpCL_cxGravarTemplate1.Enabled = true;
            CpCL_cxGravarTemplate2.Enabled = true;
            CpCL_cxGravarTemplate3.Enabled = true;
            CpCL_cxGravarTemplate4.Enabled = true;
            CpCL_cxGravarTemplate5.Enabled = true;

            btLimparLer.Enabled = true;
            btLimparGravar.Enabled = true;



            MessageBox.Show("Conectado esta funcionando corretamente", "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void btExportarFunc_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                comunicador.Parametros.CaminhoArquivo = saveFileDialog1.FileName;
                int rr = comunicador.GerarListaUsuarios();
                if (rr != CKREP3.Defs.RetornoDefs.RETORNO_OK)
                {
                    MessageBox.Show("Falha ao exportar Funcionarios.prv " + rr, "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    MessageBox.Show("Gerado com sucesso!", "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btconverter_Click(object sender, EventArgs e)
        {
            string[] templates;
            string templatesConcatenados = campoTemplateConcatenado.Text;
            if (templatesConcatenados.Length < 499)
            {
                MessageBox.Show("Template inválido");
            }
            int erro = comunicador.ConverterParaTemplateIndividual(templatesConcatenados, out templates);
            if (erro != CKREP3.Defs.RetornoDefs.RETORNO_OK)
            {
                MessageBox.Show("Erro:" + erro.ToString());
            }
            else
            {
                if (templates != null)
                {
                    cxCadTemplate1.Text = "";
                    cxCadTemplate2.Text = "";
                    cxCadTemplate3.Text = "";
                    cxCadTemplate4.Text = "";
                    cxCadTemplate5.Text = "";
                    QuantTemplatesLidos.Text = "";
                    cxCadNTemplate1.Text = "";
                    cxCadNTemplate2.Text = "";
                    cxCadNTemplate3.Text = "";
                    cxCadNTemplate4.Text = "";
                    cxCadNTemplate5.Text = "";

                    int quantTemplates = templates.Length;
                    QuantTemplatesLidos.Text = quantTemplates.ToString();
                    if (quantTemplates > 0) cxCadTemplate1.Text = templates[0];
                    if (quantTemplates > 1) cxCadTemplate2.Text = templates[1];
                    if (quantTemplates > 2) cxCadTemplate3.Text = templates[2];
                    if (quantTemplates > 3) cxCadTemplate4.Text = templates[3];
                    if (quantTemplates > 4) cxCadTemplate5.Text = templates[4];

                    //Exemplo de como pegar de qual dedo foi capturado do leitor
                    //No Rep não é possivel saber o dedo, então a numeração é dada na ordem do cadastro
                    if (quantTemplates > 0) cxCadNTemplate1.Text = Convert.FromBase64String(templates[0])[0].ToString();
                    if (quantTemplates > 1) cxCadNTemplate2.Text = Convert.FromBase64String(templates[1])[0].ToString();
                    if (quantTemplates > 2) cxCadNTemplate3.Text = Convert.FromBase64String(templates[2])[0].ToString();
                    if (quantTemplates > 3) cxCadNTemplate4.Text = Convert.FromBase64String(templates[3])[0].ToString();
                    if (quantTemplates > 4) cxCadNTemplate5.Text = Convert.FromBase64String(templates[4])[0].ToString();
                }
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        

    }
}
