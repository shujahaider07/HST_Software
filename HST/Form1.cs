using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows.Forms;

namespace HST
{
    public partial class Form1 : Form
    {
        int FinalCost;
        int snsCost;
        int voipCost;
        int expenses;
        int rent;
        int maintenance;
        int salary;
        int utilitybill;
        int internetbill;
        int bonus;
        public Form1()
        {
            InitializeComponent();

        }


        String cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;

        private void Form1_Load(object sender, EventArgs e)
        {

            rentadd();
            ExpensesCost();
            maintenanceAdd();
            saladd();
            utilityAdd();
            internetAdd();
            bonusAdd();
            

        }


        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Amjad a1 = new Amjad();
            a1.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            loan l1 = new loan();
            l1.Show();

        }

        private void totalInformationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Details d1 = new Details();
            d1.Show();
        }


        private void viewExpensesOnlyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExpensesOnly e1 = new ExpensesOnly();
            e1.Show();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            utilityAdd();
            saladd();
            rentadd();
            VoipDatagridView1();
            AllDatagridView1();
            finalCost();
            snsCostGrid();
            ExpensesCost();
            voipCostGrid();
            internetAdd();
            maintenanceAdd();
            bonusAdd();
            
        }

        private void amjadDetailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AmjadDetails am = new AmjadDetails();
            am.Show();
        }

        private void calculatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("C://Windows//System32/calc.exe");
        }



        private void pictureBox4_Click(object sender, EventArgs e)
        {
            SnsViewcs sn = new SnsViewcs();
            sn.Show();


        }


        private void pictureBox5_Click(object sender, EventArgs e)
        {
            VoipView vp = new VoipView();
            vp.Show();
        }

        private void label12_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            ExpensesOnly eo = new ExpensesOnly();
            eo.Show();

        }

        public void finalCost()
        {


            FinalCost = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                FinalCost = FinalCost + Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);


            }

            label1.Text = FinalCost.ToString();


        }
        public void snsCostGrid()
        {

            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            String qry = "Select sum(amount) from expenses where purpose = 'sns' ";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label6.Text = dr[0].ToString();

            }
            sql.Close();


        }
        //# Total Of rent , maint, salaries etc 
        //public void Addtotal()
        //{

        //    //try
        //    //{
        //    //    values = Convert.ToInt32(label32.Text);
        //    //    values1 = Convert.ToInt32(label33.Text);
        //    //    values2 = Convert.ToInt32(label34.Text);
        //    //    values3 = Convert.ToInt32(label35.Text);
        //    //    values4 = Convert.ToInt32(label36.Text);
        //    //    values5 = Convert.ToInt32(label37.Text);
        //    //    add = add + values + values1 + values2 + values3 + values4 + values5;

        //    //    label41.Text = add.ToString();
        //    //    label41.Visible = true;
        //    //}
        //    //catch (Exception)
        //    //{

        //    //}



        //}
        public void voipCostGrid()
        {


            voipCost = 0;
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                voipCost = voipCost + Convert.ToInt32(dataGridView3.Rows[i].Cells[3].Value);


            }
            label15.Text = voipCost.ToString();


        }


        public void ExpensesCost()
        {

            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            String qry = "select Sum(amount) from expenses where purpose = 'expenses' or purpose = 'data' or purpose = 'others' ";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label18.Text = dr[0].ToString();
            }

            sql.Close();



            int cost = 10000;
            int cost2 = 15000;

            if (expenses >= cost)
            {

                label18.ForeColor = System.Drawing.Color.Orange;
            }
            else if (expenses > cost2)
            {
                label18.ForeColor = System.Drawing.Color.DarkRed;
            }
            else if (expenses <= cost)
            {
                label18.ForeColor = System.Drawing.Color.Green;

            }




        }



        private void pictureBox7_Click(object sender, EventArgs e)
        {
            addData ad = new addData();
            ad.Show();
        }

        public void AllDatagridView1()
        {
            SqlConnection sql = new SqlConnection(cs);
            String qry = "select * from expenses";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;


        }


        public void VoipDatagridView1()
        {
            SqlConnection sql = new SqlConnection(cs);
            String qry = "select * from expenses where purpose = 'VOIP'";
            SqlDataAdapter da = new SqlDataAdapter(qry, sql);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView3.DataSource = dt;


        }



        private void mAILToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Mail ma = new Mail();
            ma.ShowDialog();
        }

        private void fIXEDCOSTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FixedCost fc = new FixedCost();
            fc.ShowDialog();

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            Details d = new Details();
            d.Show();
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            ExpensesOnly e1 = new ExpensesOnly();
            e1.Show();
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            SnsViewcs sns = new SnsViewcs();
            sns.Show();
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {

            VoipView v1 = new VoipView();
            v1.Show();
        }



        public void rentadd()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from rent";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label32.Text = dr[0].ToString();
            }
            sql.Close();

        }


        public void maintenanceAdd()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from maintenance";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label33.Text = dr[0].ToString();
            }
            sql.Close();
        }



        public void saladd()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from salaries";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label35.Text = dr[0].ToString();
            }
            sql.Close();
        }

        public void utilityAdd()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from utilitybill";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label34.Text = dr[0].ToString();
            }
            sql.Close();

        }
        public void internetAdd()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from internetbill1";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label36.Text = dr[0].ToString();
            }

            sql.Close();


        }

        public void bonusAdd()
        {
            SqlConnection sql = new SqlConnection(cs);
            sql.Open();
            string qry = "select sum(amount) from bonus";
            SqlCommand cmd = new SqlCommand(qry, sql);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                label37.Text = dr[0].ToString();
            }
            sql.Close();
        }

        private void bONUSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bonus b1 = new bonus();
            b1.ShowDialog();

        }

        private void pictureBox17_Click(object sender, EventArgs e)
        {
            bonusView b = new bonusView();
            b.ShowDialog();
        }
    }
}