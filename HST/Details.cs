using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace HST
{
    public partial class Details : Form
    {

        public Details()
        {
            InitializeComponent();
            textBox5.KeyUp += TextBox5_KeyUp;
            //  dataGridView1.KeyUp += DataGridView1_KeyUp;

        }


        private void TextBox5_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                BindGridView();
            }
        }

        string cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void Details_Load(object sender, EventArgs e)
        {
            BindGridView();
            style();

        }

        public void BindGridView()
        {
            SqlConnection sql = new SqlConnection(cs);
            string qry = "select * from expenses ";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
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

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                dataGridView1.CurrentRow.Selected = true;
                textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells["id"].Value.ToString();
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells["name"].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells["purpose"].Value.ToString();
                textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells["Givenby"].Value.ToString();
                dateTimePicker1.Text = dataGridView1.Rows[e.RowIndex].Cells["date"].Value.ToString();
                textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells["amount"].Value.ToString();
                textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells["Others"].Value.ToString();
            }
            catch (Exception)
            {

            }
        }

        public void EditGridData()
        {

            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "Update expenses set name = '" + textBox2.Text + "' , purpose ='" + textBox3.Text + "' , amount = '" + textBox6.Text + "' ,  Givenby = '" + textBox4.Text + "' ,  date = '" + dateTimePicker1.Value + "' where id = '" + textBox1.Text + "'";

            SqlCommand cmd = new SqlCommand(qry, sql);
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Updated item");

                BindGridView();
            }
            else
            {
                MessageBox.Show("Updated failed");
            }


            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox6.Text = "";




            sql.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            EditGridData();

        }

        public void DeleteDataGrid()
        {
            DialogResult dr = MessageBox.Show("Are you sure to Delete the date", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr == DialogResult.Yes)
            {


                SqlConnection sql = new SqlConnection(cs);
                String qry = "DELETE FROM Expenses WHERE id = '" + textBox1.Text + "'";
                SqlCommand cmd = new SqlCommand(qry, sql);
                sql.Open();
                int a = cmd.ExecuteNonQuery();
                if (a > 0)
                {
                    MessageBox.Show("Deleted");
                    BindGridView();
                }

                sql.Close();
            }
            else
            {
                MessageBox.Show(" ");
            }
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox6.Text = "";

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DeleteDataGrid();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            string qry = "SELECT * from expenses WHERE name LIKE '%" + textBox5.Text + "%' ";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void importbtn_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are You want to import Data To Excel ", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
                label9.Visible = true;
            {
                label9.Visible = true;
                progressBar1.Maximum = 100;
                progressBar1.Minimum = 0;
                progressBar1.Value = 100;


                progressBar1.Show();


                if (dataGridView1.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workbook = xcelApp.Workbooks.Add(Type.Missing);
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                    worksheet = workbook.Sheets[1];
                    worksheet = workbook.ActiveSheet;
                    worksheet.Name = "HST Expenses";

                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }


                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    try
                    {
                        xcelApp.Columns.AutoFit();
                        var saveFileDialoge = new SaveFileDialog();
                        saveFileDialoge.FileName = "DailyExpenses";
                        saveFileDialoge.DefaultExt = ".xlsx";

                        if (saveFileDialoge.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        }
                        xcelApp.Quit();

                    }
                    catch (Exception)
                    {

                    }
                }
                else
                {
                    MessageBox.Show("Not imported");
                }
            }
        }

        private void dateTimePicker3_CloseUp(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            DateTime fromDate = Convert.ToDateTime(dateTimePicker2.Value);
            DateTime toDate = Convert.ToDateTime(dateTimePicker3.Value);
            if (fromDate <= toDate)
            {
                TimeSpan ts = toDate.Subtract(fromDate);
                int days = Convert.ToInt32(ts.Days);

                String qry = "Select * from expenses where date between '" + dateTimePicker2.Value + "' and '" + dateTimePicker3.Value + "' ";
                SqlCommand cmd = new SqlCommand(qry, sql);
                sql.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(cmd.ExecuteReader());
                dataGridView1.DataSource = dt;
                sql.Close();


            }

        }

       
    }

}