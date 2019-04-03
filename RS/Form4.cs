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



namespace PUR
{
    public partial class Form2 : Form
    {
        private OleDbConnection con = new OleDbConnection(); 

        public Form2()
        {
            InitializeComponent();
            
        }
        public string mypath;
        private void Form2_Load(object sender, EventArgs e)
        {
            mypath = Directory.GetCurrentDirectory();
            pictureBox1.Image = Image.FromFile(Path.Combine(mypath, "Prosh.jpg"));
            comboBox1.Enabled = false;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            if (radioButton3.Checked == true)
            {
                try
                {
                    con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Ts-03\\users\\OGK\\Кокшаров С.А\\db\\DB.accdb;Jet OLEDB:Database Password=6567604";
                    con.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    string query = "SELECT nrvs from bdr";
                    cmd.CommandText = query;
                    OleDbDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        comboBox1.Items.Add(reader[0].ToString());
                    }
                    con.Close();
                    comboBox1.Enabled = true;
                    label23.Text = "Сетевая база данных подключена";
                    label23.ForeColor = System.Drawing.Color.Green;
                }
                catch { MessageBox.Show("База данных не доступна. Возможно нет подключения к сети.", "Ошибка"); radioButton4.Checked = true; }
            }
            if (radioButton4.Checked == true)
            {
                con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(mypath, "DB.accdb") + ";Jet OLEDB:Database Password=6567604";
                con.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                string query = "SELECT nrvs from bdr";
                cmd.CommandText = query;
                OleDbDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    comboBox1.Items.Add(reader[0].ToString());
                }
                con.Close();
                comboBox1.Enabled = true;
                label23.Text ="Локальная база данных подключена";
                label23.ForeColor = System.Drawing.Color.Green;
            }
        }

        public static String xurma(double vvmf) //функция Q из mf
        {
            int cg = 981; //g м/с^2 константа
            int cdelt = 101000; //дельта P, г/см^2 константа
            double cgam = 0.808; //гамма, г/см^3 константа
            double cq = 588; //Q, см^3/мин константа
            int ct = 60; //t, сек константа
            double pkor = Math.Sqrt(2 * cg * cdelt / cgam); //корень из 2gp/j
            double pmf = 100 * cq / (ct * pkor);
            double vmf = vvmf;
            double qson = Math.Round((vmf * ct * pkor / 100), 2); //Q, см^3/мин Расход Sonplas
            return Convert.ToString(qson);
        }

        public static String xurma2(double qqson) //функция mf из Q
        {
            int cg = 981; //g м/с^2 константа
            int cdelt = 101000; //дельта P, г/см^2 константа
            double cgam = 0.808; //гамма, г/см^3 константа
            double cq = 588; //Q, см^3/мин константа
            int ct = 60; //t, сек константа
            double pkor = Math.Sqrt(2 * cg * cdelt / cgam); //корень из 2gp/j
            double pmf = 100 * cq / (ct * pkor);
            double qson = qqson;
            double vmf = Math.Round(((qson * 100) / (pkor * ct)), 4); //Q, см^3/мин Расход Sonplas
            return Convert.ToString(vmf);
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            bool text15 = textBox15.Text.Contains(".");
            if (text15)
            {
                textBox15.Text = textBox15.Text.Replace(".", ",");
                textBox15.SelectionStart = textBox15.Text.Length;
            }
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            bool text16 = textBox16.Text.Contains(".");
            if (text16)
            {
                textBox16.Text = textBox16.Text.Replace(".", ",");
                textBox16.SelectionStart = textBox16.Text.Length;
            }
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            bool text20 = textBox20.Text.Contains(".");
            if (text20)
            {
                textBox20.Text = textBox20.Text.Replace(".", ",");
                textBox20.SelectionStart = textBox20.Text.Length;
            }
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            bool text21 = textBox21.Text.Contains(".");
            if (text21)
            {
                textBox21.Text = textBox21.Text.Replace(".", ",");
                textBox21.SelectionStart = textBox21.Text.Length;
            }
        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            bool text22 = textBox22.Text.Contains(".");
            if (text22)
            {
                textBox22.Text = textBox22.Text.Replace(".", ",");
                textBox22.SelectionStart = textBox22.Text.Length;
            }
        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {
            bool text27 = textBox27.Text.Contains(".");
            if (text27)
            {
                textBox27.Text = textBox27.Text.Replace(".", ",");
                textBox27.SelectionStart = textBox27.Text.Length;
            }
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {
            bool text28 = textBox28.Text.Contains(".");
            if (text28)
            {
                textBox28.Text = textBox28.Text.Replace(".", ",");
                textBox28.SelectionStart = textBox28.Text.Length;
            }
        }

        private void textBox29_TextChanged(object sender, EventArgs e)
        {
            bool text29 = textBox29.Text.Contains(".");
            if (text29)
            {
                textBox29.Text = textBox29.Text.Replace(".", ",");
                textBox29.SelectionStart = textBox29.Text.Length;
            }
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            bool text30 = textBox30.Text.Contains(".");
            if (text30)
            {
                textBox30.Text = textBox30.Text.Replace(".", ",");
                textBox30.SelectionStart = textBox30.Text.Length;
            }
        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {
            bool text31 = textBox31.Text.Contains(".");
            if (text31)
            {
                textBox31.Text = textBox31.Text.Replace(".", ",");
                textBox31.SelectionStart = textBox31.Text.Length;
            }
        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {
            bool text32 = textBox32.Text.Contains(".");
            if (text32)
            {
                textBox32.Text = textBox32.Text.Replace(".", ",");
                textBox32.SelectionStart = textBox32.Text.Length;
            }
        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {
            bool text33 = textBox33.Text.Contains(".");
            if (text33)
            {
                textBox33.Text = textBox3.Text.Replace(".", ",");
                textBox33.SelectionStart = textBox33.Text.Length;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Вводимые переменные

            //string ind;
            string dat = Convert.ToString(dateTimePicker1);
            string zad = textBox1.Text;
            string nrvs = textBox2.Text;
            string nkr = textBox3.Text;
            string nig = textBox4.Text;
            string hodig = textBox5.Text;
            string hodigv = textBox6.Text;
            string hodign = textBox7.Text;
            string vhvost = textBox10.Text;
            string vhvostv = textBox9.Text;
            string vhvostn = textBox8.Text;
            string prnahod = textBox13.Text;
            string prnahodv = textBox12.Text;
            string prnahodn = textBox11.Text;
            string gidropl = textBox14.Text;
            string mfv = textBox15.Text;
            string mfn = textBox16.Text;
            string zag = textBox17.Text;
            string procpolir = textBox18.Text;
            string procdrosel = textBox19.Text;
            string pmarkbosch = textBox26.Text;
            string kmarkbosch = textBox25.Text;
            string pmarkazpi = textBox24.Text;
            string kmarkazpi = textBox23.Text;
            string uvp1 = maskedTextBox1.Text;
            string uvp2 = maskedTextBox2.Text;
            string uvp3 = maskedTextBox3.Text;
            string uvp4 = maskedTextBox4.Text;
            string uvp5 = maskedTextBox5.Text;
            string uvp6 = maskedTextBox6.Text;
            string uvp7 = maskedTextBox7.Text;
            string uvp8 = maskedTextBox8.Text;
            string uvp9 = maskedTextBox9.Text;
            string uvp10 = maskedTextBox10.Text;
            string uvsh1 = maskedTextBox11.Text;
            string uvsh2 = maskedTextBox12.Text;
            string uvsh3 = maskedTextBox13.Text;
            string uvsh4 = maskedTextBox14.Text;
            string uvsh5 = maskedTextBox15.Text;
            string uvsh6 = maskedTextBox16.Text;
            string uvsh7 = maskedTextBox17.Text;
            string uvsh8 = maskedTextBox18.Text;
            string uvsh9 = maskedTextBox19.Text;
            string uvsh10 = maskedTextBox20.Text;
            double ki1 = 0;
            double ki2 = 0;
            double ki3 = 0;
            double ki4 = 0;
            double ki5 = 0;
            double ki6 = 0;
            double ki7 = 0;
            double ki8 = 0;
            double ki9 = 0;
            double ki10 = 0;
            string diamotv = textBox54.Text;

            if (textBox15.TextLength == 0 || textBox16.TextLength == 0)
            {
                MessageBox.Show("Введите mf");
                return;
            }
            //Вычисляемые переменные
            double pmfrvs1 = Convert.ToDouble(mfn);
            double pmfrvs2 = Convert.ToDouble(mfv);
            double ptrvs1 = 0.202 / pmfrvs1;
            double ptrvs2 = 0.202 / pmfrvs2;
            double pqrvs1 = Convert.ToDouble(xurma(pmfrvs1));
            double pqrvs2 = Convert.ToDouble(xurma(pmfrvs2));


            double pmfpol1 = pmfrvs1 * (Convert.ToDouble(procdrosel) / 100) + pmfrvs1;
            double pmfpol2 = pmfrvs2 * (Convert.ToDouble(procdrosel) / 100) + pmfrvs2;
            double ptpol1 = 0.202 / pmfpol1;
            double ptpol2 = 0.202 / pmfpol2;
            double pqpol1 = Convert.ToDouble(xurma(pmfpol1));
            double pqpol2 = Convert.ToDouble(xurma(pmfpol2));


            double pmfpr1 = pmfpol1 - (pmfpol1 * (Convert.ToDouble(procpolir) / 100));
            double pmfpr2 = pmfpol2 - (pmfpol2 * (Convert.ToDouble(procpolir) / 100));
            double ptpr1 = 0.202 / pmfpr1;
            double ptpr2 = 0.202 / pmfpr2;
            double pqpr1 = Convert.ToDouble(xurma(pmfpr1));
            double pqpr2 = Convert.ToDouble(xurma(pmfpr2));

            double raz = 0.176;
            double raz2 = 0.395;

            if (radioButton5.Checked)
            {
                if (textBox20.Text != "")
                {
                    ki1 = Convert.ToDouble(textBox20.Text) - raz2;
                }
                

                if (textBox21.Text != "")
                {
                    ki2 = Convert.ToDouble(textBox21.Text) - raz2;
                }
                

                if (textBox22.Text != "")
                {
                    ki3 = Convert.ToDouble(textBox22.Text) - raz2;
                }
                

                if (textBox27.Text != "")
                {
                    ki4 = Convert.ToDouble(textBox27.Text) - raz2;
                }
                

                if (textBox28.Text != "")
                {
                    ki5 = Convert.ToDouble(textBox28.Text) - raz2;
                }
                

                if (textBox29.Text != "")
                {
                    ki6 = Convert.ToDouble(textBox29.Text) - raz2;
                }
                

                if (textBox30.Text != "")
                {
                    ki7 = Convert.ToDouble(textBox30.Text) - raz2;
                }
                

                if (textBox31.Text != "")
                {
                    ki8 = Convert.ToDouble(textBox31.Text) - raz2;
                }
                

                if (textBox32.Text != "")
                {
                    ki9 = Convert.ToDouble(textBox32.Text) - raz2;
                }
                

                if (textBox33.Text != "")
                {
                    ki10 = Convert.ToDouble(textBox33.Text) - raz2;
                }
                

            }

            if (radioButton1.Checked)
            {
                if (textBox20.Text != "")
                {
                    ki1 = raz + Convert.ToDouble(textBox20.Text);
                }
                
                
                if (textBox21.Text != "")
                {
                    ki2 = raz + Convert.ToDouble(textBox21.Text);
                }
                

                if (textBox22.Text != "")
                {
                    ki3 =raz + Convert.ToDouble(textBox22.Text);
                }
                

                if (textBox27.Text != "")
                {
                    ki4 = raz + Convert.ToDouble(textBox27.Text);
                }
                

                if (textBox28.Text != "")
                {
                    ki5 = raz + Convert.ToDouble(textBox28.Text);
                }
                

                if (textBox29.Text != "")
                {
                    ki6 = raz + Convert.ToDouble(textBox29.Text);
                }
                

                if (textBox30.Text != "")
                {
                    ki7 = raz + Convert.ToDouble(textBox30.Text);
                }
                

                if (textBox31.Text != "")
                {
                    ki8 = raz + Convert.ToDouble(textBox31.Text);
                }
                

                if (textBox32.Text != "")
                {
                    ki9 = raz + Convert.ToDouble(textBox32.Text);
                }
                

                if (textBox33.Text != "")
                {
                    ki10 = raz + Convert.ToDouble(textBox33.Text);
                }
                

            }

            if (radioButton2.Checked)
            {

                if (textBox20.Text != "")
                {
                    ki1 = Convert.ToDouble(textBox20.Text);
                }
                

                if (textBox21.Text != "")
                {
                    ki2 = Convert.ToDouble(textBox21.Text);
                }
                

                if (textBox22.Text != "")
                {
                    ki3 = Convert.ToDouble(textBox22.Text);
                }
                

                if (textBox27.Text != "")
                {
                    ki4 = Convert.ToDouble(textBox27.Text);
                }
                

                if (textBox28.Text != "")
                {
                    ki5 = Convert.ToDouble(textBox28.Text);
                }
                

                if (textBox29.Text != "")
                {
                    ki6 = Convert.ToDouble(textBox29.Text);
                }
                

                if (textBox30.Text != "")
                {
                    ki7 = Convert.ToDouble(textBox30.Text);
                }
                

                if (textBox31.Text != "")
                {
                    ki8 = Convert.ToDouble(textBox31.Text);
                }
                

                if (textBox32.Text != "")
                {
                    ki9 = Convert.ToDouble(textBox32.Text);
                }
                

                if (textBox33.Text != "")
                {
                    ki10 = Convert.ToDouble(textBox33.Text);
                }
                
                               
            }
                       
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open((Path.Combine(mypath, "appur.xlsm")),  ReadOnly: false, Password: "6567604");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["variable"] as Microsoft.Office.Interop.Excel.Worksheet;

            Excel.Range userRange = x.UsedRange;
            //x.Cells[2, 2] = dat;
            x.Cells[2, 3] = zad;
            x.Cells[2, 4] = nrvs;
            x.Cells[2, 5] = nkr;
            x.Cells[2, 6] = nig;
            x.Cells[2, 7] = hodig;
            x.Cells[2, 8] = hodigv;
            x.Cells[2, 9] = hodign;
            x.Cells[2, 10] = vhvost;
            x.Cells[2, 11] = vhvostv;
            x.Cells[2, 12] = vhvostn;
            x.Cells[2, 13] = prnahod;
            x.Cells[2, 14] = prnahodv;
            x.Cells[2, 15] = prnahodn;
            x.Cells[2, 16] = gidropl;
            x.Cells[2, 17] = mfn;
            x.Cells[2, 18] = mfv;
            x.Cells[2, 19] = zag;
            x.Cells[2, 20] = procpolir;
            x.Cells[2, 21] = procdrosel;
            x.Cells[2, 22] = pmarkbosch;
            x.Cells[2, 23] = kmarkbosch;
            x.Cells[2, 24] = pmarkazpi;
            x.Cells[2, 25] = kmarkazpi;
            x.Cells[2, 26] = uvp1;
            x.Cells[2, 27] = uvp2;
            x.Cells[2, 28] = uvp3;
            x.Cells[2, 29] = uvp4;
            x.Cells[2, 30] = uvp5;
            x.Cells[2, 31] = uvp6;
            x.Cells[2, 32] = uvp7;
            x.Cells[2, 33] = uvp8;
            x.Cells[2, 34] = uvp9;
            x.Cells[2, 35] = uvp10;
            x.Cells[2, 36] = uvsh1;
            x.Cells[2, 37] = uvsh2;
            x.Cells[2, 38] = uvsh3;
            x.Cells[2, 39] = uvsh4;
            x.Cells[2, 40] = uvsh5;
            x.Cells[2, 41] = uvsh6;
            x.Cells[2, 42] = uvsh7;
            x.Cells[2, 43] = uvsh8;
            x.Cells[2, 44] = uvsh9;
            x.Cells[2, 45] = uvsh10;
            x.Cells[2, 46] = ki1;
            x.Cells[2, 47] = ki2;
            x.Cells[2, 48] = ki3;
            x.Cells[2, 49] = ki4;
            x.Cells[2, 50] = ki5;
            x.Cells[2, 51] = ki6;
            x.Cells[2, 52] = ki7;
            x.Cells[2, 53] = ki8;
            x.Cells[2, 54] = ki9;
            x.Cells[2, 55] = ki10;
            x.Cells[2, 56] = diamotv;
            x.Cells[2, 57] = pmfpr1;
            x.Cells[2, 58] = pmfpr2;
            x.Cells[2, 59] = ptpr1;
            x.Cells[2, 60] = ptpr2;
            x.Cells[2, 61] = pqpr1;
            x.Cells[2, 62] = pqpr2;
            x.Cells[2, 63] = pmfpol1;
            x.Cells[2, 64] = pmfpol2;
            x.Cells[2, 65] = ptpol1;
            x.Cells[2, 66] = ptpol2;
            x.Cells[2, 67] = pqpol1;
            x.Cells[2, 68] = pqpol2;
            x.Cells[2, 69] = pmfrvs1;
            x.Cells[2, 70] = pmfrvs2;
            x.Cells[2, 71] = ptrvs1;
            x.Cells[2, 72] = ptrvs2;
            x.Cells[2, 73] = pqrvs1;
            x.Cells[2, 74] = pqrvs2;

            Microsoft.Office.Interop.Excel.Worksheet x2 = excel.Sheets["pu"] as Microsoft.Office.Interop.Excel.Worksheet;

            Excel.Range userRange2 = x2.UsedRange;
            x2.Cells[56, 11] = textBox34.Text;   

            //Печать        
            if (checkBox1.Checked)
            {
                excel.ActiveWorkbook.PrintOutEx(17, 17);
            }

            if (checkBox2.Checked)
            {
                excel.ActiveWorkbook.PrintOutEx(12, 12);
            }

            if (checkBox3.Checked)
            {
                excel.ActiveWorkbook.PrintOutEx(13, 13);
            }

            if (checkBox4.Checked)
            {
                excel.ActiveWorkbook.PrintOutEx(10, 10);
            }

            if (checkBox8.Checked)
            {
                excel.ActiveWorkbook.PrintOutEx(11, 11);
            }

            if (checkBox7.Checked)
            {
                excel.ActiveWorkbook.PrintOutEx(14, 14);
            }

            if (checkBox6.Checked)
            {
                excel.ActiveWorkbook.PrintOutEx(18, 18);
            }

            //excel.ActiveWorkbook.PrintOutEx(4, 4);
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void данныеКонстантРасчётаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("g = 981 м/с^2 " + "\n" + "deltaP = 101000 г/с^2 " + "\n" + "gamma = 0.83 г/с^3 " + "\n" + "Q = 588 см^3/мин " + "\n" + "t = 60 сек.", "Входные данные");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = con;
            cmd.CommandText = "insert into bdr (dat, zad, nrvs, nkr, nig, hodig, hodigv, hodign, vhvost, vhvostv, vhvostn, prnahod, prnahodv, prnahodn, gidropl, mfv, mfn, zag, procpolir, procdrosel, pmarkbosch, kmarkbosch, pmarkazpi, kmarkazpi, uvp1, uvp2, uvp3, uvp4, uvp5, uvp6, uvp7, uvp8, uvp9, uvp10, uvsh1, uvsh2, uvsh3, uvsh4, uvsh5, uvsh6, uvsh7, uvsh8, uvsh9, uvsh10, ki1, ki2, ki3, ki4, ki5, ki6, ki7, ki8, ki9, ki10, diamotv) values ('" + dateTimePicker1.Text+"','"+textBox1.Text+"','" + textBox2.Text+"','" + textBox3.Text+"','" + textBox4.Text+"','" + textBox5.Text+"','" + textBox6.Text+"','" + textBox7.Text+"','" + textBox10.Text+"','" + textBox9.Text+"','" + textBox8.Text+ "','" + textBox13.Text + "','" + textBox12.Text + "','" + textBox11.Text + "','" + textBox14.Text + "','" + textBox15.Text + "','" + textBox16.Text + "','" + textBox17.Text + "','" + textBox18.Text + "','" + textBox19.Text + "','" + textBox26.Text + "','" + textBox25.Text + "','" + textBox24.Text + "','" + textBox23.Text + "', '" + maskedTextBox1.Text + "','" + maskedTextBox2.Text + "','" + maskedTextBox3.Text + "','" + maskedTextBox4.Text + "','" + maskedTextBox5.Text + "','" + maskedTextBox6.Text + "','" + maskedTextBox7.Text + "','" + maskedTextBox8.Text + "','" + maskedTextBox9.Text + "','" + maskedTextBox10.Text + "','" + maskedTextBox11.Text + "','" + maskedTextBox12.Text + "','" + maskedTextBox13.Text + "','" + maskedTextBox14.Text + "','" + maskedTextBox15.Text + "','" + maskedTextBox16.Text + "','" + maskedTextBox17.Text + "','" + maskedTextBox18.Text + "','" + maskedTextBox19.Text + "','" + maskedTextBox20.Text + "','" + textBox20.Text + "','" + textBox21.Text + "','" + textBox22.Text + "','" + textBox27.Text + "','" + textBox28.Text + "','" + textBox29.Text + "','" + textBox30.Text + "','" + textBox31.Text + "','" + textBox32.Text + "','" + textBox33.Text + "','" + textBox54.Text + "')";
            cmd.ExecuteNonQuery();
            con.Close();
            button3.PerformClick();
            if (radioButton3.Checked == true)
            { MessageBox.Show("Распылитель'" + textBox2.Text + "' добавлен в сетевую базу данных", "Успешно"); }
            if (radioButton4.Checked == true)
            { MessageBox.Show("Распылитель'" + textBox2.Text + "' добавлен в локальную базу данных", "Успешно"); }
        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = con;
            string query = "SELECT * from bdr where nrvs='" + comboBox1.Text + "'";
            cmd.CommandText = query;
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                label2.Text = reader["dat"].ToString();
                textBox1.Text = reader["zad"].ToString();
                textBox2.Text = reader["nrvs"].ToString();
                textBox3.Text = reader["nkr"].ToString();
                textBox4.Text = reader["nig"].ToString();
                textBox5.Text = reader["hodig"].ToString();
                textBox6.Text = reader["hodigv"].ToString();
                textBox7.Text = reader["hodign"].ToString();
                textBox10.Text = reader["vhvost"].ToString();
                textBox9.Text = reader["vhvostv"].ToString();
                textBox8.Text = reader["vhvostn"].ToString();
                textBox13.Text = reader["prnahod"].ToString();
                textBox12.Text = reader["prnahodv"].ToString();
                textBox11.Text = reader["prnahodn"].ToString();
                textBox14.Text = reader["gidropl"].ToString();
                textBox15.Text = reader["mfv"].ToString();
                textBox16.Text = reader["mfn"].ToString();
                textBox17.Text = reader["zag"].ToString();
                textBox18.Text = reader["procpolir"].ToString();
                textBox19.Text = reader["procdrosel"].ToString();
                textBox26.Text = reader["pmarkbosch"].ToString();
                textBox25.Text = reader["kmarkbosch"].ToString();
                textBox24.Text = reader["pmarkazpi"].ToString();
                textBox23.Text = reader["kmarkazpi"].ToString();
                maskedTextBox1.Text = reader["uvp1"].ToString();
                maskedTextBox2.Text = reader["uvp2"].ToString();
                maskedTextBox3.Text = reader["uvp3"].ToString();
                maskedTextBox4.Text = reader["uvp4"].ToString();
                maskedTextBox5.Text = reader["uvp5"].ToString();
                maskedTextBox6.Text = reader["uvp6"].ToString();
                maskedTextBox7.Text = reader["uvp7"].ToString();
                maskedTextBox8.Text = reader["uvp8"].ToString();
                maskedTextBox9.Text = reader["uvp9"].ToString();
                maskedTextBox10.Text = reader["uvp10"].ToString();
                maskedTextBox11.Text = reader["uvsh1"].ToString();
                maskedTextBox12.Text = reader["uvsh2"].ToString();
                maskedTextBox13.Text = reader["uvsh3"].ToString();
                maskedTextBox14.Text = reader["uvsh4"].ToString();
                maskedTextBox15.Text = reader["uvsh5"].ToString();
                maskedTextBox16.Text = reader["uvsh6"].ToString();
                maskedTextBox17.Text = reader["uvsh7"].ToString();
                maskedTextBox18.Text = reader["uvsh8"].ToString();
                maskedTextBox19.Text = reader["uvsh9"].ToString();
                maskedTextBox20.Text = reader["uvsh10"].ToString();
                textBox20.Text = reader["ki1"].ToString();
                textBox21.Text = reader["ki2"].ToString();
                textBox22.Text = reader["ki3"].ToString();
                textBox27.Text = reader["ki4"].ToString();
                textBox28.Text = reader["ki5"].ToString();
                textBox29.Text = reader["ki6"].ToString();
                textBox30.Text = reader["ki7"].ToString();
                textBox31.Text = reader["ki8"].ToString();
                textBox32.Text = reader["ki9"].ToString();
                textBox33.Text = reader["ki10"].ToString();
                textBox54.Text = reader["diamotv"].ToString();


            }
            con.Close();
        }

        

        private void хочешьБытьСчастливымToolStripMenuItem_Click(object sender, EventArgs e)
        {

           
            MessageBox.Show("Будь им.  ","Азаза лалка");
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("LDP(лень двигатель прогресса:-)" + "\n" + "Программа для создания планов управления " + "\n" + "опытных распылителей на серийном оборудовании." + "\n" + "v1.0 - калёная линия с прошивки." + "\n" + "©Кокшаров Сергей Александрович", "LDP");
        }

        private void очистиьВсеФормыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label2.Text = "";
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox10.Clear();
            textBox9.Clear();
            textBox8.Clear();
            textBox13.Clear();
            textBox12.Clear();
            textBox11.Clear();
            textBox14.Clear();
            textBox15.Clear();
            textBox16.Clear();
            textBox17.Clear();
            textBox18.Clear();
            textBox19.Clear();
            textBox26.Clear();
            textBox25.Clear();
            textBox24.Clear();
            textBox23.Clear();
            maskedTextBox1.Clear();
            maskedTextBox2.Clear();
            maskedTextBox3.Clear();
            maskedTextBox4.Clear();
            maskedTextBox5.Clear();
            maskedTextBox6.Clear();
            maskedTextBox7.Clear();
            maskedTextBox8.Clear();
            maskedTextBox9.Clear();
            maskedTextBox10.Clear();
            maskedTextBox11.Clear();
            maskedTextBox12.Clear();
            maskedTextBox13.Clear();
            maskedTextBox14.Clear();
            maskedTextBox15.Clear();
            maskedTextBox16.Clear();
            maskedTextBox17.Clear();
            maskedTextBox18.Clear();
            maskedTextBox19.Clear();
            maskedTextBox20.Clear();
            textBox20.Clear();
            textBox21.Clear();
            textBox22.Clear();
            textBox27.Clear();
            textBox28.Clear();
            textBox29.Clear();
            textBox30.Clear();
            textBox31.Clear();
            textBox32.Clear();
            textBox33.Clear();
            textBox54.Clear();
            textBox34.Clear();
        }

        private void создатьBackupСетевойToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string put1 = @"\\Ts-03\users\OGK\Кокшаров С.А\db\DB.accdb";
            string put2 = @"\\Ts-03\users\OGK\Кокшаров С.А\db\backup\DB_"+DateTime.Now.Date.ToString("dd.MM.yyyy")+".accdb";
            File.Copy(put1,put2, true);
        }

        private void обновитьЛокальнуюИзСетевойToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Файл локальной базы данных будет заменён на файл сетевой базы данных и восстановить его будет невозможно. Ты уверен?", "Обновление локальной базы данных", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string put1 = @"\\Ts-03\users\OGK\Кокшаров С.А\db\DB.accdb";
                string put2 = Path.Combine(mypath, "DB.accdb");
                File.Copy(put1, put2, true);
            }
        }
    }

}