using System;
using System.Data.OleDb;
using System.Windows.Forms;


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



            /*float km, peage, repas, hebergement;

            km = float.Parse(textBox4.Text);
            peage = float.Parse(textBox5.Text);
            repas = float.Parse(textBox6.Text);
            hebergement = float.Parse(textBox7.Text);
            */
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "INSERT INTO frais ([date],[motif],[trajet],[km],[cout-peage],[cout-repas],[cout-hebergement],[adresse-mail]) VALUES ('"+textBox1.Text+"','"+comboBox1.Text+"','"+textBox3.Text+ "','" +textBox4.Text+ "','" +textBox5.Text+ "','" +textBox6.Text+ "','"+textBox7.Text+"','kijiramer@hotmail.fr')";


                command.ExecuteNonQuery();
                MessageBox.Show("Ajout de frais pris en compte");
                connection.Close();
            }catch(Exception ex)
            {
                MessageBox.Show("Erreur " + ex);
            }

            this.Close();

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
    }
}
