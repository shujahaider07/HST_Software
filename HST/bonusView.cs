using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;


namespace HST
{
    public partial class bonusView : Form
    {
        public bonusView()
        {
            InitializeComponent();
        }
        string cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
        private void bonusView_Load(object sender, EventArgs e)
        {
            View();
        }
    
    public void View()
        {

            SqlConnection sql = new SqlConnection(cs);
            string qry = "select amount , date from bonus";
            SqlDataAdapter da = new SqlDataAdapter(qry,sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;


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

                String qry = "Select amount , date from bonus where date between '" + dateTimePicker2.Value + "' and '" + dateTimePicker3.Value + "' ";
                SqlCommand cmd = new SqlCommand(qry, sql);
                sql.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(cmd.ExecuteReader());
                dataGridView1.DataSource = dt;
                sql.Close();


            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {

                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = xcelApp.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                worksheet = workbook.Sheets[1];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "HST bONUS";

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
                saveFileDialoge.FileName = "bonus";
                saveFileDialoge.DefaultExt = ".xlsx";

                if (saveFileDialoge.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                }
                xcelApp.Quit();

            }
        }
    }
}
