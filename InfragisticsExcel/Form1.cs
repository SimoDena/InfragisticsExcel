using Infragistics.Documents.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InfragisticsExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: questa riga di codice carica i dati nella tabella 'dataSet1.Custumer'. È possibile spostarla o rimuoverla se necessario.
            this.custumerTableAdapter.Fill(this.dataSet1.Custumer);
        }

        private void salvaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.ultraGridExcelExporter1.Export(this.ultraGrid1,"custumers.xls");
        }

        private void modificaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Workbook wb = Workbook.Load("custumers.xls");
            Random rnd = new Random();

            for (int i = 1; i < 20; i++)
            {
                for (int k = 0; k < 3; k++)
                {
                    if (i%2==0)
                    {
                        wb.Worksheets[0].Rows[i].Cells[1].Value = "Mario";
                        wb.Worksheets[0].Rows[i].Cells[2].Value = rnd.Next(1,3);
                        wb.Worksheets[0].Rows[i].Cells[0].Value = i;
                        wb.Worksheets[0].Rows[i].Cells[k].CellFormat.FillPatternBackgroundColor = Color.DarkRed;
                    }
                    else
                    {
                        wb.Worksheets[0].Rows[i].Cells[1].Value = "Luigi";
                        wb.Worksheets[0].Rows[i].Cells[2].Value = rnd.Next(1, 4);
                        wb.Worksheets[0].Rows[i].Cells[0].Value = i;
                        wb.Worksheets[0].Rows[i].Cells[k].CellFormat.FillPatternBackgroundColor = Color.YellowGreen;
                    }
                }
            }

            wb.Save("custumersModificato.xls");
        }

        private void caricaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Workbook wb = Workbook.Load("custumers.xls");
            DataTable dtCustumers = new DataTable("CLIENTI");

            DataColumn colonna;
            DataRow riga;

            colonna = new DataColumn();
            colonna.DataType= Type.GetType("System.Int32"); 
            colonna.ColumnName = "idCliente";
            dtCustumers.Columns.Add(colonna);

            colonna = new DataColumn();
            colonna.DataType = Type.GetType("System.String");
            colonna.ColumnName = "nome";
            dtCustumers.Columns.Add(colonna);

            colonna = new DataColumn();
            colonna.DataType = Type.GetType("System.Int32");
            colonna.ColumnName = "macchina";
            colonna.AllowDBNull = true;
            dtCustumers.Columns.Add(colonna);

            for (int i = 1; i < wb.Worksheets[0].Rows.Count(); i++)
            {
                riga = dtCustumers.NewRow();
                riga["idCliente"] = wb.Worksheets[0].Rows[i].Cells[0].Value;
                riga["nome"] = wb.Worksheets[0].Rows[i].Cells[1].Value;
                riga["macchina"] = wb.Worksheets[0].Rows[i].Cells[2].Value;
                dtCustumers.Rows.Add(riga);              
            }
            dataSet1.Tables.Add(dtCustumers);
            ultraGrid1.DataSource = dtCustumers;
            ultraGrid1.Refresh();
        }
    }
}
