using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
namespace HST
{
    public partial class addData : Form
    {

        public addData()
        {
            InitializeComponent();
            purposetxt.KeyUp += Purposetxt_KeyUp;
        }

        private void Purposetxt_KeyUp(object sender, KeyEventArgs e)
        {

            if (purposetxt.SelectedItem == "Others")
            {
                richTextBox1.Visible = true;

            }
            else if (purposetxt.CanSelect == true)
            {
                richTextBox1.Visible = false;
            }

        }

        string cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void button1_Click(object sender, EventArgs e)
        {
            AddData();

        }
        public void AddData()
        {

            SqlConnection sql = new SqlConnection(cs);
            String qry = "insert into Expenses values(@name , @purpose , @amount  , @givenby , @date)";
            SqlCommand cmd = new SqlCommand(qry, sql);
            cmd.Parameters.AddWithValue("@name", nametxt.Text);
            cmd.Parameters.AddWithValue("@purpose", purposetxt.Text);
            cmd.Parameters.AddWithValue("@amount", amounttxt.Text);
            cmd.Parameters.AddWithValue("@givenby", givenbytxt.Text);
            cmd.Parameters.AddWithValue("@date", dateTimePicker1.Value);
            sql.Open();
            int a = cmd.ExecuteNonQuery();

            if (a > 0)
            {
                MessageBox.Show("sucessfull	✔");
            }
            else
            {
                MessageBox.Show("failed ☒");
            }
            sql.Close();

            nametxt.Text = "";
            purposetxt.Text = "";
            amounttxt.Text = "";
            givenbytxt.Text = "";
            nametxt.Focus();


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

        public void style()
        {

            dataGridView1.BorderStyle = BorderStyle.None;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkSlateGray;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.White;
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("MS Reference Sans serif", 10);
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;


        }

        public void bindGridView()
        {
            SqlConnection sql = new SqlConnection(cs);
            String qry = "select name,amount , purpose , givenby , date  from expenses";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void addData_Load(object sender, EventArgs e)
        {
            purposetxt.Text = "Expenses";
            String[] arr = { "Expenses", "Voip", "Sns", "Data" ,"Others" };
            foreach (String item in arr)
            {
                purposetxt.Items.Add(item);
            }


            style();
            bindGridView();
            panel2.Visible = true;
        }


        public void countall()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select count (id) from expenses";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {

                DataCount.Text = reader[0].ToString();


            }
        }

        private void addData_Activated(object sender, EventArgs e)
        {
            countall();
            bindGridView();
        }


    }


}
