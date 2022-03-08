using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;

namespace HST
{
    public partial class loan : Form
    {
        public loan()
        {
            InitializeComponent();
        }
        String cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void loan_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "Cash";
            bindGridView();
            String[] arr = { "Cash" , "Cheque"};
            foreach (var item in arr)
            {
                comboBox1.Items.Add(item);
            }
            style();

        }
        public void insertData()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "insert into loan values (@name , @discription , @amount , @From , @cheque , @Remaining, @date)";

            SqlCommand cmd = new SqlCommand(qry,sql);
         
            cmd.Parameters.AddWithValue("@name",nametxt.Text);
            cmd.Parameters.AddWithValue("@discription",descripttxt.Text);
            cmd.Parameters.AddWithValue("@amount",amnttxt.Text);
            cmd.Parameters.AddWithValue("@From",frmtxt.Text);
            cmd.Parameters.AddWithValue("@cheque",comboBox1.SelectedItem);
            cmd.Parameters.AddWithValue("@Remaining",remainingtxt.Text);
            cmd.Parameters.AddWithValue("@date",dateTimePicker1.Value);     
           
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("Data Inserted");
                bindGridView();
            }
            else
            {
                MessageBox.Show("Failed");
            }

            clearData();
            sql.Close();

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
        public void clearData()
        {
            nametxt.Text = "";
            descripttxt.Text = "";
            amnttxt.Text = "";
            frmtxt.Text = "";
            remainingtxt.Text = "";
            
        }
        public void bindGridView()
        {
            SqlConnection sql = new SqlConnection(cs);
            string qry = "select * from loan";
            SqlDataAdapter da = new SqlDataAdapter(qry,sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            insertData();
        }

        private void button2_Click(object sender, EventArgs e)
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
    }
}
