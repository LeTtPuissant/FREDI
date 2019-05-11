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
                string query = "SELECT * FROM [frais]";
                command.CommandText = query;

                OleDbDataAdapter da = new OleDbDataAdapter(command);
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur " + ex);
            }

            

            textBox3.Text = "Salle d'Armes de Villers lès Nancy, 1 rue Rodin - 54600 Villers lès Nancy";
            

            // Autocomplésion pour le formulaire note de frais...

            /* OleDbCommand command1 = new OleDbCommand();
            command1.Connection = connection;
            string query1 = "select nom, prenom from Demandeurs";
            command1.CommandText = query1;

            

            OleDbDataReader reader = command1.ExecuteReader();
            while (reader.Read())
            {
                string[] souss = { reader.["nom"], reader.["prenom"] };
                textBox1.Text = souss;
                textBox2.Text = reader["prenom"].ToString();
                textBox3.Text = reader["Date de naissance"].ToString();
                textBox4.Text = reader["rue"].ToString();
                textBox5.Text = reader["cp"].ToString();
                textBox6.Text = reader["ville"].ToString();

            } */


            connection.Close();
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Ajout ajout = new Ajout();
            ajout.Show();
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

            // Soussigné
            xlWorkSheet.get_Range("A5", "H5").HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A5", "H5").VerticalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A5", "H5").Font.Size = 10;
            xlWorkSheet.get_Range("A5", "H5").Value = textBox1.Text;

            // Deumerant
            xlWorkSheet.get_Range("A7", "H7").HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A7", "H7").VerticalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A7", "H7").Font.Size = 10;
            xlWorkSheet.get_Range("A7", "H7").Value = textBox2.Text;

            // Année civile
            xlWorkSheet.get_Range("A7", "H7").HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A7", "H7").VerticalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A7", "H7").Font.Size = 16;
            xlWorkSheet.get_Range("F2", "H2").Value = "Année civile " + textBox8.Text;

            

            //  Total
            xlWorkSheet.get_Range("I17", "I17").HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("I17", "I17").VerticalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("I17", "I17").Font.Size = 10;
            xlWorkSheet.get_Range("I17", "I17").Value = textBox4.Text + "€";

            // Reprentant légal de 
            xlWorkSheet.get_Range("A20", "H21").HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A20", "H21").VerticalAlignment = excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("A20", "H21").Font.Size = 10;
            xlWorkSheet.get_Range("A20", "H21").Value = richTextBox1.Text;


            


            xlWorkBook.SaveAs("Bordereau de note de frais.xlsx", XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
            xlWorkBook.Close(0);
            xlApp.Quit();

            


            MessageBox.Show("Votre feuille de note de frais est enregistré sur votre Bureau");
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void RichTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            
        }
    }
}
    
