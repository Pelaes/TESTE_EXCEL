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
using System.IO;
//using Microsoft.Office.Tools.Excel;//fazer referencia no assembly
//using Microsoft.Office.Interop.Excel;//fazer referencia no assembly

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
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();  //cria objeto para buscar arquivo
                openFileDialog1.Filter = "Arquivos Excel (*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb";//cria filtros para arquivos excel
                openFileDialog1.FilterIndex = 3;

                openFileDialog1.Multiselect = false;        //valor falso não permite seleção multipla de arquivos
                openFileDialog1.Title = "Escolha um arquivo Microsoft Excel";   //define o nome da caixa de dialogo
                openFileDialog1.InitialDirectory = @"Desktop"; //define o diretorio inicial

                if (openFileDialog1.ShowDialog() == DialogResult.OK)//Abrindo um arquivo
                {
                    string pathName = openFileDialog1.FileName; //retorna a string com o endereço do arquivo
                    //string fileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName); //retorna a string com o nome do arquivo sem extensão
                    DataTable tbContainer = new DataTable(); // cria um objeto datatable
                    string strConn = string.Empty;
                    //string sheetName = filename;

                    FileInfo file = new FileInfo(pathName);
                    if (!file.Exists)
                    {
                        throw new Exception("Error, file doesn't exists!");
                    }
                    string extension = file.Extension;
                    switch (extension)
                    {
                        case ".xls":
                            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                        case ".xlsx":
                            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                            break;
                        default:
                            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                    }
                    OleDbConnection cnnxls = new OleDbConnection(strConn);
                    OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [planilha1$]"), cnnxls);
                    oda.Fill(tbContainer);

                    dataGridView1.DataSource = tbContainer;
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }
        }
    }
}
