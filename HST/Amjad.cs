using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace HST
{
    public partial class Amjad : Form
    {
        public Amjad()
        {
            InitializeComponent();
            amnttxt.KeyUp += Amnttxt_KeyUp;
            comboBox1.KeyUp += ComboBox1_KeyUp;
            dateTimePicker1.KeyUp += DateTimePicker1_KeyUp;
        }

        private void DateTimePicker1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Enter)
            {
                button1.Focus();
            }
        }

        
        private void ComboBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Enter)
            {
                monthlysalarytxt.Focus();
            }
        }

        private void Amnttxt_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Enter)
            {
                comboBox1.Focus();
            }
        }

        String cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void Amjad_Load(object sender, EventArgs e)
        {
            bindGridView();
            string[] arr = { "YES", "NO" };
            foreach (var item in arr)
            {
                comboBox1.Items.Add(item);
            }


        }
        public void bindGridView()
        {
            SqlConnection sql = new SqlConnection(cs);
            string qry = "select * from amjad";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "insert into amjad values (@amount , @Give , @date , @monthlysalary,@purpose)";
            SqlCommand cmd = new SqlCommand(qry, sql);
            cmd.Parameters.AddWithValue("@amount", amnttxt.Text);
            cmd.Parameters.AddWithValue("@Give", comboBox1.SelectedItem);
            cmd.Parameters.AddWithValue("@date", dateTimePicker1.Value);
            cmd.Parameters.AddWithValue("@monthlysalary", monthlysalarytxt.Text);
            cmd.Parameters.AddWithValue("@purpose", purposetxt.Text);
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Inserted");
                bindGridView();
            }
            else
            {
                MessageBox.Show("Failed");
            }
            sql.Close();
            amnttxt.Text = "";
            monthlysalarytxt.Text = "";
            purposetxt.Text = "";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView1.CurrentRow.Selected = true;
                Idtxt.Text = dataGridView1.Rows[e.RowIndex].Cells["id"].Value.ToString();
                amnttxt.Text = dataGridView1.Rows[e.RowIndex].Cells["amount"].Value.ToString();
                comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells["give"].Value.ToString();
                dateTimePicker1.Text = dataGridView1.Rows[e.RowIndex].Cells["date"].Value.ToString();
                monthlysalarytxt.Text = dataGridView1.Rows[e.RowIndex].Cells["monthlysalary"].Value.ToString();
              purposetxt.Text = dataGridView1.Rows[e.RowIndex].Cells["purpose"].Value.ToString();
                
            }
            catch (Exception)
            {

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DeleteDataGrid();
        }
        public void DeleteDataGrid()
        {
            DialogResult dr =  MessageBox.Show("Are you want to Delete the Data","Alert",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);

            if (dr == DialogResult.Yes)
            {


                SqlConnection sql = new SqlConnection(cs);
                String qry = "DELETE FROM amjad WHERE id = '" + Idtxt.Text + "'";
                SqlCommand cmd = new SqlCommand(qry, sql);
                sql.Open();
                int a = cmd.ExecuteNonQuery();
                if (a > 0)
                {
                    MessageBox.Show("Deleted");
                    bindGridView();
                }

                sql.Close();
            }
            else
            {
                MessageBox.Show("Not Deleted");
            }
            amnttxt.Text = "";
            givetxt.Text = "";
            monthlysalarytxt.Text = "";
            Idtxt.Text = "";
            purposetxt.Text = "";         
        }

        private void amnttxt_KeyPress(object sender, KeyPressEventArgs e)
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

        private void monthlysalarytxt_KeyPress(object sender, KeyPressEventArgs e)
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
