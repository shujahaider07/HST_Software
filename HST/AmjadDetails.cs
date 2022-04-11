using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace HST
{
    public partial class AmjadDetails : Form
    {
        public AmjadDetails()
        {
            InitializeComponent();
            dataGridView1.KeyUp += DataGridView1_KeyUp;
        }

        private void DataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2)
            {
                panel2.Visible = true;

            }
            else if (e.KeyCode == Keys.Escape)
            {
                panel2.Visible = false;

            }
        }

        String cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;

        private void AmjadDetails_Load(object sender, EventArgs e)
        {
            String[] arr = { "YES", "NO" };
            foreach (var item in arr)
            {
                comboBox1.Items.Add(item);

            }

            style();


            BindGridview();




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
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("MS Reference Sans serif", 10);
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;


        }
        public void BindGridview()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            String qry = "select * from amjad";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;


        }


        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are you sure to Delete the Data", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr == DialogResult.Yes)
            {
                SqlConnection sql = new SqlConnection(cs);
                String qry = "DELETE FROM amjad WHERE id = '" + idtxt.Text + "'";
                SqlCommand cmd = new SqlCommand(qry, sql);
                sql.Open();
                int a = cmd.ExecuteNonQuery();
                if (a > 0)
                {
                    MessageBox.Show("Deleted");
                    BindGridview();
                }
                else
                {
                    MessageBox.Show("Not Deleted");
                }
                sql.Close();

            }
            else
            {
                MessageBox.Show("Not Deleted");
            }

            idtxt.Text = "";
            amonttxt.Text = "";
            monthlytxt.Text = "";
            purposetxt.Text = "";

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                dataGridView1.CurrentRow.Selected = true;
                idtxt.Text = dataGridView1.Rows[e.RowIndex].Cells["id"].Value.ToString();
                amonttxt.Text = dataGridView1.Rows[e.RowIndex].Cells["amount"].Value.ToString();
                comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells["give"].Value.ToString();
                dateTimePicker1.Text = dataGridView1.Rows[e.RowIndex].Cells["date"].Value.ToString();
                monthlytxt.Text = dataGridView1.Rows[e.RowIndex].Cells["monthlysalary"].Value.ToString();
                purposetxt.Text = dataGridView1.Rows[e.RowIndex].Cells["purpose"].Value.ToString();

            }
            catch (Exception)
            {

            }


        }
        public void EditGridView()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "Update amjad set amount = '" + amonttxt.Text + "' , give ='" + comboBox1.SelectedItem + "' , date = '" + dateTimePicker1.Value + "' , monthlySalary = '" + monthlytxt.Text + "' , purpose = '"+purposetxt.Text+"' where id = '" + idtxt.Text + "'";

            SqlCommand cmd = new SqlCommand(qry, sql);
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Updated item");
                BindGridview();
            }
            else
            {
                MessageBox.Show("Updated failed");
            }

            idtxt.Text = "";
            amonttxt.Text = "";
            monthlytxt.Text = "";
            purposetxt.Text = "";




            sql.Close();

        }


        private void button1_Click(object sender, EventArgs e)
        {
            EditGridView();
        }

        private void button3_Click(object sender, EventArgs e)
        {

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

        }

        private void dateTimePicker3_CloseUp(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            DateTime toTime = Convert.ToDateTime(dateTimePicker2.Value);
            DateTime fromTime = Convert.ToDateTime(dateTimePicker3.Value);
            if (fromTime <= toTime)
            {

                TimeSpan ts = toTime.Subtract(fromTime);
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

