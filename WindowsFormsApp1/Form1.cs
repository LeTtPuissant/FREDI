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
    public partial class Form1 : Form
    {
        private OleDbConnection connection = new OleDbConnection();

        public Form1()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\kijiramer\Desktop\M2L-Projet-FDERI\frais.mdb";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txt_user.Text = "admin";
            txt_password.Text = "root";
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void clickk(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            command.CommandText = "select * from Adherents where Nom='" + txt_user.Text + "' and Prenom='" + txt_password.Text + "' ";
            OleDbDataReader reader = command.ExecuteReader();

            int count = 0;

            while (reader.Read())
            {
                count++;
            }
            if (count == 1)
            {
                
                Information Info = new Information();
                Info.Show();
                Hide();

            }
            else if (count > 1)
            {
                MessageBox.Show("erreur de doublon");
            }
            else
            {
                MessageBox.Show("Identifiant ou Mot de passe invalide réessayer");
            }
            connection.Close();

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Inscription Inscri = new Inscription();
            Inscri.Show();
            this.Hide();


        }

        private void Txt_user_TextChanged(object sender, EventArgs e)
        {

        }

        private void Txt_password_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
