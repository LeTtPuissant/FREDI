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
    public partial class Information : Form
    {
        private OleDbConnection connection = new OleDbConnection();
        public Information()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\kijiramer\Desktop\M2L-Projet-FDERI\frais.mdb";
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            MainForm noteform = new MainForm();
            noteform.Show();
            Hide();
        }

        private void Information_Load_1(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            string query = "select * from Demandeurs";
            command.CommandText = query;

            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                textBox1.Text = reader["nom"].ToString();
                textBox2.Text = reader["prenom"].ToString();
                textBox3.Text = reader["Date de naissance"].ToString();
                textBox4.Text = reader["rue"].ToString();
                textBox5.Text = reader["cp"].ToString();
                textBox6.Text = reader["ville"].ToString();
               
            }
            connection.Close();
        }

        private void Button2_Click(object sender, EventArgs e)
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
    }
}
