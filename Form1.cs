using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using objExcel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Examen
{
    public partial class Form1 : Form
    {
        public string Path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        public string Path2 = @"C:\\Users\\ADMIN2020\\OneDrive\\Escritorio\\Ciudadanos.xlsx";
        List<Ciudadanos> ciudadanos = new List<Ciudadanos>();
        List<Partidos> Partidos = new List<Partidos>();
        static int V;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument(Path2);
            int iRow = 2;
            int mayor;

            int MAYOR = 0;
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
            {
                Ciudadanos op = new Ciudadanos();
                op.DPI1 = sl.GetCellValueAsInt32(iRow, 1);
                op.Nombre1 = sl.GetCellValueAsString(iRow, 2);
                comboBox1.Items.Add(sl.GetCellValueAsString(iRow,2));
                ciudadanos.Add(op);
                iRow++;
            }
    }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Partidos;

            objExcel.Application objAplicacion = new objExcel.Application();
            Workbook objLibro = objAplicacion.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet objHoja = (Worksheet)objAplicacion.ActiveSheet;

            objAplicacion.Visible = false;
            foreach (DataGridViewColumn columna in dataGridView1.Columns)
            {
                objHoja.Cells[1, columna.Index + 1] = columna.HeaderText;
                foreach (DataGridViewRow fila in dataGridView1.Rows)
                {
                    objHoja.Cells[fila.Index + 2, columna.Index + 1] = fila.Cells[columna.Index].Value;

                }
            }
            objLibro.SaveAs(Path + "\\PartidosPoliticos.xlsx");
            objLibro.Close();
            objAplicacion.Quit();
        }

        public void Mostrar()
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = Partidos;
            dataGridView1.Refresh();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Partidos op = new Partidos();
            Total ad= new Total();

            op.Nombre1=comboBox1.SelectedItem.ToString();
            op.NombreDelPartido1=textBox1.Text;
            op.Fecha1=DateTime.Now;
            V++;
            Partidos.Add(op);


            comboBox1.SelectedIndex=0;
            textBox1.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Partidos=Partidos.OrderBy(a => a.NombreDelPartido1).ToList();
            Mostrar();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            label4.Text=V.ToString();
        }
    }
}
