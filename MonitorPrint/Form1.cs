using System;
using System.Windows.Forms;
using System.Printing;
using System.IO;
using System.Xml;

namespace MonitorPrint
{
    public partial class Form1 : Form
    {
        private const string ARQUIVO = @"C:\Users\erik_\Documents\Disco E\log";
        int ANTERIOR;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "INICIAR")
            {
                timer1.Interval = 500;
                timer1.Start();
                button1.Text = "PARAR";
            }
            else
            {
                timer1.Stop();
                button1.Text = "INICIAR";
            }
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            PrintServer SERVIDOR = new PrintServer();
            PrintQueueCollection IMPRESSORAS = SERVIDOR.GetPrintQueues();

            foreach(PrintQueue IMPRESSORA in IMPRESSORAS)
            {
                int PAGINAS = 0;
                int IDENTIFICADOR = 0;
                string DOCUMENTO = "";
                try
                {
                    if(IMPRESSORA.NumberOfJobs > 0)
                    {
                        IMPRESSORA.Refresh();
                        PrintJobInfoCollection IMPRESSOES = IMPRESSORA.GetPrintJobInfoCollection();

                        foreach (var IMPRESSAO in IMPRESSOES)
                        {

                            PAGINAS = IMPRESSAO.NumberOfPages;
                            IDENTIFICADOR = IMPRESSAO.JobIdentifier;
                            DOCUMENTO = IMPRESSAO.Name;
                        }
                    }
                }
                catch(Exception) { }

                if((PAGINAS > 0) && (IDENTIFICADOR != ANTERIOR))
                {
                    ListViewItem LINHA = new ListViewItem(DateTime.Now.ToString());
                     LINHA.SubItems.Add(PAGINAS.ToString());
                     LINHA.SubItems.Add(IMPRESSORA.Name);
                     LINHA.SubItems.Add(Environment.UserName);
                     LINHA.SubItems.Add(Environment.MachineName);
                     listView1.Items.Add(LINHA);
                     ANTERIOR = IDENTIFICADOR;

                    // Cria o nome do arquivo com ano, mês, dia, hora minuto e segundo

                    string nomeArquivo = ARQUIVO + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xml";

                    // Cria um novo arquivo e devolve um StreamWriter para ele

                    StreamWriter writer = new StreamWriter(nomeArquivo);

                    // Agora é só sair escrevendo
                    writer.WriteLine(DateTime.Now.ToString());
                    writer.WriteLine(IMPRESSORA.Name);
                    writer.WriteLine(PAGINAS);
                    writer.WriteLine(Environment.UserName);
                    writer.WriteLine(Environment.MachineName);
                    writer.WriteLine(DOCUMENTO);
                    writer.WriteLine(IDENTIFICADOR);


                    // Não esqueça de fechar o arquivo ao terminar

                    writer.Close();

                }
            }

        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Button1_Paint(object sender, PaintEventArgs e)
        {

        }

     }
}
