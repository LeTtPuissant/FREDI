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
using excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;


namespace WindowsFormsApp1
{
    public partial class MainForm : Form
    {
        private OleDbConnection connection = new OleDbConnection();

        public MainForm()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\kijiramer\Desktop\M2L-Projet-FDERI\frais.mdb";
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label5_Click(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Voulez vous quitter?",
            "Quitter", MessageBoxButtons.YesNo);

            if (dialog == DialogResult.Yes)
            {
                Form1 form1 = new Form1();
                form1.Show();
                Hide();
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                string query = "SELECT * FROM [Lignes-frais]";
                command.CommandText = query;

                OleDbDataAdapter da = new OleDbDataAdapter(command);
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;

                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur " + ex);
            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void Button2_Click(object sender, EventArgs e)
        {

        }

        private void Button6_Click(object sender, EventArgs e)
        {

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Button5_Click(object sender, EventArgs e)
        {
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            excel.Application xlApp = new excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlApp.Sheets.Add(Type: @"C:\Users\kijiramer\Desktop\M2L-Projet-FDERI\Bordereau de note de frais-1");

            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[3, 7] = "Supinfo";
            xlWorkSheet.Cells[3, 6] = "test gros";

            xlWorkBook.SaveAs("supinfo.xlsx", excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, excel.XlSaveAsAccessMode.xlNoChange, excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
            xlWorkBook.Close(0);
            xlApp.Quit();


            MessageBox.Show("Ecriture check");
        }
            
    }
}
    
