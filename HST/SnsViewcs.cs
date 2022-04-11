using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;


namespace HST
{
    public partial class SnsViewcs : Form
    {
        public SnsViewcs()
        {
            InitializeComponent();
        }

        string cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void SnsViewcs_Load(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            String qry = "select * from expenses where purpose = 'sns' or purpose = 'Others' ";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            dataGridView1.DataSource = dt;
            da.Fill(dt);
            style();


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


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "SELECT * from expenses WHERE purpose = 'sns' and name LIKE '%" + textBox1.Text + "%' ";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;



            sql.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DialogResult dr = MessageBox.Show("Are You want to import Data To Excel ", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)

            {

                if (dataGridView1.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workbook = xcelApp.Workbooks.Add(Type.Missing);
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                    worksheet = workbook.Sheets[1];
                    worksheet = workbook.ActiveSheet;
                    worksheet.Name = "HST Sns";

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
                        saveFileDialoge.FileName = "SNS";
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

                String qry = "Select * from expenses where purpose = 'sns' and date between '" + dateTimePicker2.Value + "' and '" + dateTimePicker3.Value + "' ";
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

