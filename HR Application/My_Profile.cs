using Bytescout.Spreadsheet;
using LiveCharts;
using LiveCharts.Defaults;
using LiveCharts.Wpf;
using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace HR_Application
{
    public partial class My_Profile : Form
    {
        int row = Form1.row;
        public My_Profile()
        {
            InitializeComponent();
            readExcel();
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
            Application.Exit();
        }

        private void guna2Button9_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.Show();
            Visible = false;
        }

        private void My_Profile_Load(object sender, EventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            My_Evaluation f3 = new My_Evaluation();
            f3.Show();
            Visible = false;
        }

        private void guna2Button4_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Letters f3 = new Letters();
            f3.Show();
            Visible = false;
        }

        private void guna2CirclePictureBox1_Click(object sender, EventArgs e)
        {

        }
        public void readExcel()
        {


            string filepath = $"{System.IO.Path.GetDirectoryName(Application.ExecutablePath)}\\test5.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filepath);
            ws = wb.Worksheets[1];
        
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
