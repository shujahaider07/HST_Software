using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace HST
{
    public partial class bonus : Form
    {
        public bonus()
        {
            InitializeComponent();
        }

        String cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void bonus_Load(object sender, EventArgs e)
        {
            BindgridView();


        }

        public void insertData()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "insert into bonus values (@amount , @date)";
            SqlCommand cmd = new SqlCommand(qry,sql);
            cmd.Parameters.AddWithValue("@amount",amounttxt.Text);
            cmd.Parameters.AddWithValue("@date",dateTimePicker1.Value);
            int a = cmd.ExecuteNonQuery();
            if (a > 0) {
                MessageBox.Show("Data Inserted");
            }
            else
            {
                MessageBox.Show("failed to insert");
            }

            sql.Close();
        }
    public void BindgridView()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();

            String qry = "select id ,amount , date from bonus ";
            SqlDataAdapter da = new SqlDataAdapter(qry,sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;


            sql.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            insertData();
        }

        private void bonus_Activated(object sender, EventArgs e)
        {
            BindgridView();
        }

       

        private void button2_Click(object sender, EventArgs e)
        {

            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "Update bonus set amount = '" + amounttxt.Text + "' , date ='" + dateTimePicker1.Value + "' where id = '" + idtxt.Text + "'";

            SqlCommand cmd = new SqlCommand(qry, sql);
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Updated item");
                BindgridView();
            }
            else
            {
                MessageBox.Show("Updated failed");
            }

            amounttxt.Text = " ";





            sql.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are you want to Delete the Data", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (dr == DialogResult.Yes)
            {


                SqlConnection sql = new SqlConnection(cs);
                String qry = "DELETE FROM bonus WHERE id = '" + idtxt.Text + "'";
                SqlCommand cmd = new SqlCommand(qry, sql);
                sql.Open();
                int a = cmd.ExecuteNonQuery();
                if (a > 0)
                {
                    MessageBox.Show("Deleted");
                    BindgridView();
                }

                sql.Close();
            }
            else
            {
                MessageBox.Show("Not Deleted");
            }
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView1.CurrentRow.Selected = true;
                idtxt.Text = dataGridView1.Rows[e.RowIndex].Cells["id"].Value.ToString();
                amounttxt.Text = dataGridView1.Rows[e.RowIndex].Cells["amount"].Value.ToString();
                dateTimePicker1.Text = dataGridView1.Rows[e.RowIndex].Cells["date"].Value.ToString();
            }
            catch (Exception)
            {

            }
        }

        private void amounttxt_KeyPress(object sender, KeyPressEventArgs e)
        {

            char ch = e.KeyChar;
            if (char.IsDigit(ch))
            {
                e.Handled = false;

            }
            else if (e.KeyChar == 8)
            {
                e.Handled = false;

            }
            else
            {

                e.Handled = true;

            }
        }
    }
}
