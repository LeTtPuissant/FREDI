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
    public partial class Inscription : Form
    {
        private OleDbConnection connection = new OleDbConnection();
        public Inscription()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\frais.mdb";
        }

        private void Inscription_Load(object sender, EventArgs e)
        {
           
        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void Label6_Click(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            Hide();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
           
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "INSERT INTO Demandeurs ([Nom], [Prenom], [Date de naissance], [rue], [cp], [ville], [adresse-mail]) VALUES ('" + textBox1.Text + "','" + textBox2.Text + "' ,'"+(textBox3.Text)+"' , '" + textBox4.Text + "', '" + textBox5.Text + "', '" + textBox6.Text + "', '"+ textBox8.Text +"') ";
                command.ExecuteNonQuery();
                MessageBox.Show("Inscription prise en compte");
                connection.Close();

            MainForm mainForm = new MainForm();
            mainForm.Show();
            Hide();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
