
using System;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;

namespace HR_Application
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button7_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button8_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2Button9_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {

            My_Profile f2 = new My_Profile();
            f2.Show();
            Visible = false;

        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            My_Evaluation f3 = new My_Evaluation();
            f3.Show();
            Visible = false;
        }

        private void guna2Button4_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Letters f3 = new Letters();
            f3.Show();
            Visible = false;
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void guna2Button10_Click(object sender, EventArgs e)
        {
            readExcel();
        }

        private void readExcel()
        {

            
            string filepath = $"{System.IO.Path.GetDirectoryName(Application.ExecutablePath)}\\test5.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filepath);
            ws = wb.Worksheets[1];
            string name = ListName.Text;
            int row;
            for (row=1;row<=ws.Rows.Count;row++)
            {
                if (ws.Cells[row, 1].Value == name)
                    break;
            }
       
            label23.Text = Convert.ToString(ws.Cells[row, 1].Value);
            label22.Text = Convert.ToString(ws.Cells[row, 2].Value);
            label4.Text = Convert.ToString(ws.Cells[row, 3].Value) + "%";
            label7.Text = Convert.ToString(ws.Cells[row, 4].Value) + "%";
            label10.Text = Convert.ToString(ws.Cells[row, 5].Value) + "%";
            label13.Text = Convert.ToString(ws.Cells[row, 6].Value) + "%";
            label20.Text = Convert.ToString(ws.Cells[row, 7].Value);

            

        }

       
    }
}
