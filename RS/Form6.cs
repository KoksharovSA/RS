using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.IO;
using System.Data.OleDb;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace PUR2
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public string mypath;
        Form1 name = new Form1();
        string dat;
        private void Form2_Load(object sender, EventArgs e)
        {
            mypath = Directory.GetCurrentDirectory();
            dat = name.dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy") ;
            label7.Text = "Внесите данные";
            this.label7.ForeColor = System.Drawing.Color.Black;
        }

        private void button1_Click(object sender, EventArgs e)     
        {
            try
            {
                label7.Text = "Жди.....";
                this.label7.ForeColor = System.Drawing.Color.Orange;
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Path.Combine(mypath, @"pu\" + textBox1.Text + textBox2.Text + " " + dat + ".xlsx"), ReadOnly: false);

                Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["mk7d"] as Microsoft.Office.Interop.Excel.Worksheet;
                Excel.Range userRange = x.UsedRange;
                if (checkBox1.Checked == true)
                { x.Cells[9, 2] = "*"; }
                else { x.Cells[9, 2] = " "; }

                if (checkBox2.Checked == true)
                { x.Cells[13, 2] = "*"; }
                else { x.Cells[13, 2] = " "; }

                if (checkBox3.Checked == true)
                { x.Cells[22, 2] = "*"; }
                else { x.Cells[22, 2] = " "; }

                if (checkBox4.Checked == true)
                { x.Cells[32, 2] = "*"; }
                else { x.Cells[32, 2] = " "; }

                if (checkBox5.Checked == true)
                { x.Cells[35, 2] = "*"; }
                else { x.Cells[35, 2] = " "; }

                if (checkBox6.Checked == true)
                { x.Cells[49, 2] = "*"; }
                else { x.Cells[49, 2] = " "; }

                if (checkBox7.Checked == true)
                { x.Cells[64, 2] = "*"; }
                else { x.Cells[64, 2] = " "; }

                if (checkBox8.Checked == true)
                { x.Cells[71, 2] = "*"; }
                else { x.Cells[71, 2] = " "; }

                if (checkBox9.Checked == true)
                { x.Cells[81, 2] = "*"; }
                else { x.Cells[81, 2] = " "; }

                if (checkBox10.Checked == true)
                { x.Cells[93, 2] = "*"; }
                else { x.Cells[93, 2] = " "; }

                x.Cells[56, 7] = textBox3.Text;
                x.Cells[115, 7] = textBox3.Text;
                x.Cells[100, 7] = textBox5.Text;

                Microsoft.Office.Interop.Excel.Worksheet x2 = excel.Sheets["mk9d"] as Microsoft.Office.Interop.Excel.Worksheet;
                Excel.Range userRange2 = x2.UsedRange;
                if (checkBox1.Checked == true)
                { x2.Cells[9, 2] = "*"; }
                else { x2.Cells[9, 2] = " "; }

                if (checkBox2.Checked == true)
                { x.Cells[13, 2] = "*"; }
                else { x.Cells[13, 2] = " "; }

                if (checkBox3.Checked == true)
                { x2.Cells[22, 2] = "*"; }
                else { x2.Cells[22, 2] = " "; }

                if (checkBox4.Checked == true)
                { x2.Cells[32, 2] = "*"; }
                else { x2.Cells[32, 2] = " "; }

                if (checkBox5.Checked == true)
                { x2.Cells[35, 2] = "*"; }
                else { x2.Cells[35, 2] = " "; }

                if (checkBox6.Checked == true)
                { x2.Cells[49, 2] = "*"; }
                else { x2.Cells[49, 2] = " "; }

                if (checkBox7.Checked == true)
                { x2.Cells[64, 2] = "*"; }
                else { x2.Cells[64, 2] = " "; }

                if (checkBox8.Checked == true)
                { x2.Cells[71, 2] = "*"; }
                else { x2.Cells[71, 2] = " "; }

                if (checkBox9.Checked == true)
                { x2.Cells[81, 2] = "*"; }
                else { x2.Cells[81, 2] = " "; }

                if (checkBox10.Checked == true)
                { x2.Cells[93, 2] = "*"; }
                else { x2.Cells[93, 2] = " "; }

                x2.Cells[56, 7] = textBox3.Text;
                x2.Cells[115, 7] = textBox3.Text;
                x2.Cells[92, 7] = textBox5.Text;

                Microsoft.Office.Interop.Excel.Worksheet x3 = excel.Sheets["mk7"] as Microsoft.Office.Interop.Excel.Worksheet;
                Excel.Range userRange3 = x3.UsedRange;
                if (checkBox1.Checked == true)
                { x3.Cells[9, 2] = "*"; }
                else { x3.Cells[9, 2] = " "; }

                if (checkBox2.Checked == true)
                { x3.Cells[13, 2] = "*"; }
                else { x3.Cells[13, 2] = " "; }

                if (checkBox3.Checked == true)
                { x3.Cells[22, 2] = "*"; }
                else { x3.Cells[22, 2] = " "; }

                if (checkBox4.Checked == true)
                { x3.Cells[32, 2] = "*"; }
                else { x3.Cells[32, 2] = " "; }

                if (checkBox5.Checked == true)
                { x3.Cells[35, 2] = "*"; }
                else { x3.Cells[35, 2] = " "; }

                if (checkBox6.Checked == true)
                { x3.Cells[49, 2] = "*"; }
                else { x3.Cells[49, 2] = " "; }

                if (checkBox7.Checked == true)
                { x3.Cells[64, 2] = "*"; }
                else { x3.Cells[64, 2] = " "; }

                if (checkBox8.Checked == true)
                { x3.Cells[71, 2] = "*"; }
                else { x3.Cells[71, 2] = " "; }

                if (checkBox9.Checked == true)
                { x3.Cells[81, 2] = "*"; }
                else { x3.Cells[81, 2] = " "; }

                if (checkBox10.Checked == true)
                { x3.Cells[93, 2] = "*"; }
                else { x3.Cells[93, 2] = " "; }

                x3.Cells[56, 7] = textBox3.Text;
                x3.Cells[115, 7] = textBox3.Text;
                x3.Cells[100, 7] = textBox5.Text;

                Microsoft.Office.Interop.Excel.Worksheet x4 = excel.Sheets["mk9"] as Microsoft.Office.Interop.Excel.Worksheet;
                Excel.Range userRange4 = x4.UsedRange;
                if (checkBox1.Checked == true)
                { x4.Cells[9, 2] = "*"; }
                else { x4.Cells[9, 2] = " "; }

                if (checkBox2.Checked == true)
                { x4.Cells[13, 2] = "*"; }
                else { x4.Cells[13, 2] = " "; }

                if (checkBox3.Checked == true)
                { x4.Cells[22, 2] = "*"; }
                else { x4.Cells[22, 2] = " "; }

                if (checkBox4.Checked == true)
                { x4.Cells[32, 2] = "*"; }
                else { x4.Cells[32, 2] = " "; }

                if (checkBox5.Checked == true)
                { x4.Cells[35, 2] = "*"; }
                else { x4.Cells[35, 2] = " "; }

                if (checkBox6.Checked == true)
                { x4.Cells[49, 2] = "*"; }
                else { x4.Cells[49, 2] = " "; }

                if (checkBox8.Checked == true)
                { x4.Cells[64, 2] = "*"; }
                else { x4.Cells[64, 2] = " "; }

                if (checkBox9.Checked == true)
                { x4.Cells[74, 2] = "*"; }
                else { x4.Cells[74, 2] = " "; }

                if (checkBox10.Checked == true)
                { x4.Cells[85, 2] = "*"; }
                else { x4.Cells[85, 2] = " "; }

                x4.Cells[56, 7] = textBox3.Text;
                x4.Cells[115, 7] = textBox3.Text;
                x4.Cells[92, 7] = textBox5.Text;

                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                label7.Text = "Данные добавлены";
                this.label7.ForeColor = System.Drawing.Color.Green;

            }
            catch { MessageBox.Show("Что-то пошло не так как задумывалось, может название файла не то или фвйл ПУ не создан?", "Ошибонька"); }
        }
    }
}
