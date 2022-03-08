using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Windows.Forms;
namespace HST
{
    public partial class Login : Form
    {
        String cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        public Login()
        {
            InitializeComponent();
            textBox1.KeyUp += TextBox1_KeyUp;
            textBox2.KeyUp += TextBox2_KeyUp;
        }

        private void TextBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.Focus();
            }
        }

        private void TextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox2.Focus();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select * from login where username = @user and password = @pass";
            SqlCommand cmd = new SqlCommand(qry, sql);
            cmd.Parameters.AddWithValue("@user", textBox1.Text);
            cmd.Parameters.AddWithValue("@pass", textBox2.Text);
            SqlDataReader dr = cmd.ExecuteReader();
            if (dr.HasRows == true)
            {
                MessageBox.Show("Login Sucessfull ✔ ");
                // DialogResult DR = MessageBox.Show("Login Sucessfull","Alert",Me);
                Form1 frm = new Form1();
                frm.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Failed to Login ☒ ");
            }



            sql.Close();

            textBox1.Text = "";
            textBox2.Text = "";




        }

        private void Login_Load(object sender, EventArgs e)
        {

        }
    }
}
