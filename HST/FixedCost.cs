using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace HST
{


    public partial class FixedCost : Form
    {
        int officeCost;
        int maintenanceCost;
        int UtilityCost;
        int salaryCost;
        int internetCost;

        public FixedCost()
        {
            InitializeComponent();
        }
        string cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void FixedCost_Load(object sender, EventArgs e)
        {
            officeRent();
            maintenance();
            internet();
            salary();
            utility();


        }




        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void FixedCost_Activated(object sender, EventArgs e)
        {
            officeRent();
            maintenance();
            internet();
            salary();
            utility();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            rent r1 = new rent();
            r1.ShowDialog();



        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            maintenance m1 = new maintenance();
            m1.ShowDialog();


        }



        private void pictureBox3_Click(object sender, EventArgs e)
        {
            utilitybill ut = new utilitybill();
            ut.ShowDialog();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            salaries s1 = new salaries();
            s1.ShowDialog();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            internetbill i = new internetbill();
            i.Show();
        }

        public void officeRent()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from rent ";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label1.Text = dr[0].ToString();
            }


            sql.Close();


        }



        public void maintenance()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from maintenance ";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label2.Text = dr[0].ToString();
            }

        }


        public void utility()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select  sum(amount) from utilitybill ";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label3.Text = dr[0].ToString();
            }
        }


        public void salary()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select  sum(amount) from salaries";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label4.Text = dr[0].ToString();
            }

        }


        public void internet()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select  sum(amount) from internetbill1 ";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label5.Text = dr[0].ToString();
            }

        }


    }

}