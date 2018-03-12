using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Carregar
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog arquivo = new OpenFileDialog();
            arquivo.Filter = "Arquivos Excel (*.xls; *.xlsx) | *.xls; *.xlsx"; //filtro para excel
            arquivo.Multiselect = false; //permite ou não seleção multipla
            arquivo.Title = "Selecione o arquivo"; //titulo da caixa de dialogo
                                                   
             //if (arquivo.ShowDialog() == DialogResult.OK)
             
             
            
             
             
            string nomesemextensao;
            if (arquivo.ShowDialog() == DialogResult.OK)
            {
                string caminho = arquivo.FileName;
                nomesemextensao = System.IO.Path.GetFileNameWithoutExtension(arquivo.FileName);
            
                string pathconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + caminho + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                OleDbConnection conn = new OleDbConnection(pathconn);
                OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + nomesemextensao + "$]", conn);
                DataTable dt = new DataTable();
                MyDataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
        }
    }
}
