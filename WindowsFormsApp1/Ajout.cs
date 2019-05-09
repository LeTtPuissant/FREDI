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


namespace WindowsFormsApp1
{
    public partial class Ajout : Form
    {
        private OleDbConnection connection = new OleDbConnection();
        public Ajout()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\kijiramer\Desktop\M2L-Projet-FDERI\frais.mdb";

        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label8_Click(object sender, EventArgs e)
        {

        }

        private void TextBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void Ajout_Load(object sender, EventArgs e)
        {

        }

        private void Label10_Click(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "INSERT INTO Lignes-frais ([date], [motif], [trajet], [km], [cout-peage], [cout-repas], [cout-hebergement]) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "' ,'" + textBox3.Text + "' , '" + textBox4.Text + "', '" + textBox5.Text + "', '" + textBox6.Text + "', '" + textBox7.Text + "') ";
                command.ExecuteNonQuery();
                MessageBox.Show("Aout de frais prise en compte");
                connection.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("erreur " + ex);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
