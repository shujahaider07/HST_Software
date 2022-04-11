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
using System.Configuration;
using System.Data.SqlClient;

namespace HST
{
    public partial class utilitybill : Form
    {
        public utilitybill()
        {
            InitializeComponent();
        }
        String cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void button6_Click(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            String qry = "insert into utilitybill values (@amount , @date, @purpose)";
            SqlCommand cmd = new SqlCommand(qry, sql);
            cmd.Parameters.AddWithValue("@amount", textBox1.Text);
            cmd.Parameters.AddWithValue("@date", dateTimePicker2.Value);
            cmd.Parameters.AddWithValue("@purpose", textBox2.Text);


            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Data Inserted");
            }
            else
            {
                MessageBox.Show("Failed to insert Data");
            }

            sql.Close();
        }

        private void utilitybill_Load(object sender, EventArgs e)
        {
            textBox2.Text = " Utilitybill";
            BindGridViewRent();
            style();
        }
        public void BindGridViewRent()
        {
            SqlConnection sql = new SqlConnection(cs);
            String qry = "Select * from  utilitybill ";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.DataSource = dt;


    
        }

        public void style()
        {

            dataGridView2.BorderStyle = BorderStyle.None;
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
            dataGridView2.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridView2.DefaultCellStyle.SelectionBackColor = Color.DarkSlateGray;
            dataGridView2.DefaultCellStyle.SelectionForeColor = Color.White;
            dataGridView2.BackgroundColor = Color.White;
            dataGridView2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("MS Reference Sans serif", 10);
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;


        }

        private void button5_Click(object sender, EventArgs e)
        {
            EditInfo();
        }
        public void EditInfo()
        {
            SqlConnection sql = new SqlConnection(cs);
            string qry = "Update utilitybill set amount = '" + textBox1.Text + "' , date ='" + dateTimePicker2.Value + "' where id = '" + textBox3.Text + "'";
            sql.Open();
            SqlCommand cmd = new SqlCommand(qry, sql);
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Updated");
                textBox1.Text = "";

            }
            else
            {
                MessageBox.Show("Update Failed");
            }
            sql.Close();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            String qry = "delete from utilitybill where id = '" + textBox3.Text + "'";
            sql.Open();
            SqlCommand cmd = new SqlCommand(qry, sql);
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Deleted");
            }
            else
            {
                MessageBox.Show("fault");
            }
            sql.Close();
        }

        private void utilitybill_Activated(object sender, EventArgs e)
        {
            BindGridViewRent();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textBox3.Text = dataGridView2.Rows[e.RowIndex].Cells["id"].Value.ToString();
                dataGridView2.CurrentRow.Selected = true;
                textBox1.Text = dataGridView2.Rows[e.RowIndex].Cells["amount"].Value.ToString();
                dateTimePicker2.Text = dataGridView2.Rows[e.RowIndex].Cells["date"].Value.ToString();
            }
            catch (Exception)
            {

            }
        }

        private void importbtn_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are You want to import Data To Excel ", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
     
            {
              


                if (dataGridView2.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workbook = xcelApp.Workbooks.Add(Type.Missing);
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                    worksheet = workbook.Sheets[1];
                    worksheet = workbook.ActiveSheet;
                    worksheet.Name = "HST UtilityBill";

                    for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                    {
                        worksheet.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                    }


                    for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView2.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    try
                    {
                        xcelApp.Columns.AutoFit();
                        var saveFileDialoge = new SaveFileDialog();
                        saveFileDialoge.FileName = "UtilityBill";
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

                String qry = "Select * from utilitybill where date between '" + dateTimePicker2.Value + "' and '" + dateTimePicker3.Value + "' ";
                SqlCommand cmd = new SqlCommand(qry, sql);
                sql.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(cmd.ExecuteReader());
                dataGridView2.DataSource = dt;
                sql.Close();


            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
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

