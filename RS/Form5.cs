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
    public partial class Form1 : Form
    {
        private OleDbConnection con = new OleDbConnection();
        public string mypath;

        
        public Form1()
        {
            InitializeComponent();

            
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            mypath = Directory.GetCurrentDirectory();
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button7.Enabled = false;
            label18.Text = "Создайте ПУ";
            this.label18.ForeColor = System.Drawing.Color.Black;

        }
        
        private void button1_Click(object sender, EventArgs e)//Создание файла ПУ
        {

            try
            {
                label18.Text = "Жди.....";
                this.label18.ForeColor = System.Drawing.Color.Orange;
                //Основные переменные
                string nas = textBox14.Text;
                string ind = textBox15.Text;
                string raz = textBox17.Text;
                string prov = textBox16.Text;
                string g = textBox1.Text;//Габаритный размер
                string k2 = textBox2.Text;//Размер до калибра 2
                string k22 = Convert.ToString(Convert.ToDouble(k2) - 0.176);//Размер до калибра 2,2
                string dku90 = textBox3.Text;//До кромки угла 90
                string k2d = textBox4.Text;//От калибра 2 до дна колодца
                string k22d = Convert.ToString(Convert.ToDouble(k2d) + 0.176);//От калибра 2,2 до дна колодца
                string ts = textBox5.Text;//Толщина стенки
                string rn = textBox7.Text;//Радиус носика
                string sn = textBox8.Text;//Скругление плечей и носика
                string k625n = textBox6.Text;//От клибра 6,25 до носика
                string gtp = textBox9.Text;//Глубина топливоподвода
                string dtp = textBox11.Text;//Диаметр топливоподвода
                string gk = textBox10.Text;//Глубина до кармана
                string dk = textBox13.Text;//Диаметр колодца без дорна
                string rd = textBox12.Text;//Радиус дорна

                //Припуски
                double prnamicron = 0.15;
                double prnapulem = 0.15;
                double prnasfer = 0.3;
                double prnaradnos = 0.3;
                double prnauva = 0.07;
                double diamkolpoddorn = 0.7;

                try
                {
                    con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Ts-03\\users\\OGK\\Кокшаров С.А\\db\\DB.accdb;Jet OLEDB:Database Password=6567604";
                    con.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    string query = "SELECT prnamicron, prnapulem, prnasfer, prnaradnos, prnauva, diamkolpoddorn from pripuski";
                    cmd.CommandText = query;
                    OleDbDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        prnamicron = Convert.ToDouble(reader[0]);
                        prnapulem = Convert.ToDouble(reader[1]);
                        prnasfer = Convert.ToDouble(reader[2]);
                        prnaradnos = Convert.ToDouble(reader[3]);
                        prnauva = Convert.ToDouble(reader[4]);
                        diamkolpoddorn = Convert.ToDouble(reader[5]);
                    }
                    con.Close();
                }
                catch (Exception){ MessageBox.Show("Ошибка","Ошибка"); }

                //Rika
                double grika = Convert.ToDouble(g) + prnamicron + prnapulem + prnasfer;//Габаритный размер для рики
                double dk85 = Convert.ToDouble(grika) - (Convert.ToDouble(k625n) + 0.41);//От торца до калибра 8,5
                double dkr94 = Convert.ToDouble(k625n) + 0.56;//От края носика до кромки диаметра 9,4
                double rnr = Convert.ToDouble(rn) + prnaradnos;//Радиус носика Rika
                double snr = Convert.ToDouble(sn);//Скругление плечей и носика Rika

                //Micron
                double dk22m = Convert.ToDouble(k22) + prnapulem - 0.02;//От торца до калибра 2,2 Micron
                double dk34m = Convert.ToDouble(dk22m) - 0.81;//От торца до калибра 3,4(0,81 разница между калибрами) Micron
                double ddkz = Convert.ToDouble(dk22m) + 1.46;//От торца до дна колодца заготовки  (детали под дорн)
                double ddkd = Convert.ToDouble(k22d) + prnauva + dk22m;//От торца до дна колодца заготовки  (детали без дорна)
                double k22ddk = Convert.ToDouble(k22d) + prnauva;//От калибра 2,2 до дна колодца  (детали без дорна)
                double dkmd = diamkolpoddorn;//Диаметр колодца с дорном
                double dkm = Convert.ToDouble(dk);//Диаметр колодца без дорна

                //Дорнование
                double k22ddkdorn = Convert.ToDouble(k22d) + prnauva;//От калибра 2,2 до дна колодца заготовки  (детали под дорн)
                double rddorn = Convert.ToDouble(rd);//Радиус дорна

                //Топляк штифты
                double gtps = Convert.ToDouble(gtp) + prnapulem;//Глубина топливоподвода
                double dtps = Convert.ToDouble(dtp);//Диаметр топливоподвода

                //Карман
                double gke = Convert.ToDouble(gk) + prnapulem;//Глубина кармана
                                
                //
                //Вставка в документ
                //

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open((Path.Combine(mypath, @"mkrasp.xlsx")), ReadOnly: false);
                
                //МК 7-ка дорн
                Microsoft.Office.Interop.Excel.Worksheet m1 = excel.Sheets["mk7d"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRangem1 = m1.UsedRange;

                m1.Cells[1, 8] = nas;
                m1.Cells[2, 8] = ind;
                m1.Cells[57, 8] = raz;
                m1.Cells[58, 8] = prov;
                m1.Cells[57, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
                m1.Cells[58, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");

                //МК 9-ка дорн            
                Microsoft.Office.Interop.Excel.Worksheet m2 = excel.Sheets["mk9d"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRangem2 = m2.UsedRange;

                m2.Cells[1, 8] = nas;
                m2.Cells[2, 8] = ind;
                m2.Cells[57, 8] = raz;
                m2.Cells[58, 8] = prov;
                m2.Cells[57, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
                m2.Cells[58, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");

                //МК 7-ка без дорна
                Microsoft.Office.Interop.Excel.Worksheet m3 = excel.Sheets["mk7"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRangem3 = m3.UsedRange;

                m3.Cells[1, 8] = nas;
                m3.Cells[2, 8] = ind;
                m3.Cells[57, 8] = raz;
                m3.Cells[58, 8] = prov;
                m3.Cells[57, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
                m3.Cells[58, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");

                //МК 9-ка без дорна
                Microsoft.Office.Interop.Excel.Worksheet m4 = excel.Sheets["mk9"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRangem4 = m4.UsedRange;

                m4.Cells[1, 8] = nas;
                m4.Cells[2, 8] = ind;
                m4.Cells[57, 8] = raz;
                m4.Cells[58, 8] = prov;
                m4.Cells[57, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
                m4.Cells[58, 10] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");

                //010
                Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["010"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange = x.UsedRange;

                x.Cells[58, 1] = nas;
                x.Cells[59, 1] = ind;
                x.Cells[58, 15] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
                x.Cells[59, 15] = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
                x.Cells[58, 11] = raz;
                x.Cells[59, 11] = prov;
                
                //020-1(С углом 30)
                Microsoft.Office.Interop.Excel.Worksheet x1 = excel.Sheets["020-1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange1 = x1.UsedRange;

                x1.Cells[14, 11] = dkr94;
                x1.Cells[16, 11] = snr;
                x1.Cells[26, 11] = rnr;
                x1.Cells[28, 4] = dk85;

                //020-1(Без угла 30)
                Microsoft.Office.Interop.Excel.Worksheet x2 = excel.Sheets["020-2"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange2 = x2.UsedRange;

                x2.Cells[14, 11] = dkr94;
                x2.Cells[16, 11] = snr;
                x2.Cells[26, 11] = rnr;
                x2.Cells[28, 4] = dk85;

                //030-1(Большая выточка под штангу, заготовка под дорн)
                Microsoft.Office.Interop.Excel.Worksheet x3 = excel.Sheets["030-1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange3 = x3.UsedRange;

                x3.Cells[18, 1] = dkmd;
                x3.Cells[23, 4] = dk34m;
                x3.Cells[24, 6] = dk22m;
                x3.Cells[26, 6] = ddkz;

                //030-2(Большая выточка под штангу, заготовка без дорна)
                Microsoft.Office.Interop.Excel.Worksheet x4 = excel.Sheets["030-2"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange4 = x4.UsedRange;

                x4.Cells[18, 1] = dkm;
                x4.Cells[23, 4] = dk34m;
                x4.Cells[24, 6] = dk22m;
                x4.Cells[26, 6] = ddkz;
                x4.Cells[11, 2] = k22ddk;

                //030-3(Маленькая выточка, заготовка под дорн)
                Microsoft.Office.Interop.Excel.Worksheet x5 = excel.Sheets["030-3"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange5 = x5.UsedRange;

                x5.Cells[18, 1] = dkmd;
                x5.Cells[23, 4] = dk34m;
                x5.Cells[24, 6] = dk22m;
                x5.Cells[26, 6] = ddkz;

                //030-4(Маленькая выточка, заготовка без дорна)
                Microsoft.Office.Interop.Excel.Worksheet x6 = excel.Sheets["030-4"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange6 = x6.UsedRange;

                x6.Cells[18, 1] = dkm;
                x6.Cells[23, 4] = dk34m;
                x6.Cells[24, 6] = dk22m;
                x6.Cells[26, 6] = ddkz;
                x6.Cells[11, 2] = k22ddk;

                //041(Дорн)
                Microsoft.Office.Interop.Excel.Worksheet x11 = excel.Sheets["041"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange11 = x11.UsedRange;

                x11.Cells[20, 13] = k22ddkdorn;
                x11.Cells[9, 11] = rd;

                //045(Топливоподвод)
                Microsoft.Office.Interop.Excel.Worksheet x12 = excel.Sheets["045"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange12 = x12.UsedRange;

                x12.Cells[11, 2] = gtps;
                x12.Cells[12, 4] = dtps;

                //110(Карман)
                Microsoft.Office.Interop.Excel.Worksheet x13 = excel.Sheets["110"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange13 = x13.UsedRange;

                x13.Cells[25, 3] = gke;
               
                //135-1(Большая выточка 7-ка)
                Microsoft.Office.Interop.Excel.Worksheet x7 = excel.Sheets["135-1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange7 = x7.UsedRange;
                if (rddorn == 0) {x7.Cells[16, 1] = "Ø"+dk; }
                else{x7.Cells[16, 1] = "R"+rddorn; }
                x7.Cells[22, 2] = k22ddk;
                x7.Cells[21, 4] = dk34m;
                x7.Cells[22, 5] = dk22m;
                x7.Cells[8, 9] = gke;
                x7.Cells[15, 11] = dtps;

                //135-2(Большая выточка 9-ка)
                Microsoft.Office.Interop.Excel.Worksheet x8 = excel.Sheets["135-2"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange8 = x8.UsedRange;
                if (rddorn == 0) { x8.Cells[16, 1] = "Ø" + dk; }
                else { x8.Cells[16, 1] = "R" + rddorn; }
                x8.Cells[22, 2] = k22ddk;
                x8.Cells[21, 4] = dk34m;
                x8.Cells[22, 5] = dk22m;
                x8.Cells[8, 9] = gke;
                x8.Cells[15, 11] = dtps;

                //135-3(Маленькая выточка 7-ка)
                Microsoft.Office.Interop.Excel.Worksheet x9 = excel.Sheets["135-3"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange9 = x9.UsedRange;
                if (rddorn == 0) { x9.Cells[16, 1] = "Ø" + dk; }
                else { x9.Cells[16, 1] = "R" + rddorn; }
                x9.Cells[22, 2] = k22ddk;
                x9.Cells[21, 4] = dk34m;
                x9.Cells[22, 5] = dk22m;
                x9.Cells[8, 9] = gke;
                x9.Cells[15, 11] = dtps;

                //135-4(Маленькая выточка 9-ка)
                Microsoft.Office.Interop.Excel.Worksheet x10 = excel.Sheets["135-4"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange10 = x10.UsedRange;
                if (rddorn == 0) { x10.Cells[16, 1] = "Ø" + dk; }
                else { x10.Cells[16, 1] = "R" + rddorn; }
                x10.Cells[22, 2] = k22ddk;
                x10.Cells[21, 4] = dk34m;
                x10.Cells[22, 5] = dk22m;
                x10.Cells[8, 9] = gke;
                x10.Cells[15, 11] = dtps;

                //175 (Пулемет)
                Microsoft.Office.Interop.Excel.Worksheet x14 = excel.Sheets["175"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange14 = x14.UsedRange;
                x14.Cells[18, 7] = Convert.ToDouble(dk22m) -0.15;

                //200-1 (Сфера 30гр)
                Microsoft.Office.Interop.Excel.Worksheet x15 = excel.Sheets["200-1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange15 = x15.UsedRange;
                x15.Cells[21, 6] = Convert.ToDouble(g) - 0.02;
                x15.Cells[21, 13] = sn;
                x15.Cells[19, 13] = rn;
                x15.Cells[10, 9] = ts;
                x15.Cells[12, 9] = k625n;

                //200-2 (Сфера без 30гр)
                Microsoft.Office.Interop.Excel.Worksheet x16 = excel.Sheets["200-2"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRange16 = x16.UsedRange;
                x16.Cells[21, 6] = Convert.ToDouble(g) - 0.02;
                x16.Cells[21, 13] = sn;
                x16.Cells[19, 13] = rn;
                x16.Cells[10, 9] = ts;
                x16.Cells[12, 9] = k625n;


                sheet.Close(true, (Path.Combine(mypath, @"pu\" + nas + ind + " " + dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy") + ".xlsx")), Type.Missing);
                excel.Quit();
                
                //Вставка наладка Rika30
                Microsoft.Office.Interop.Excel.Application excel2 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet2 = excel2.Workbooks.Open((Path.Combine(mypath, @"rika30\Rika30.xlsx")), ReadOnly: false);
                Microsoft.Office.Interop.Excel.Worksheet r1 = excel2.Sheets["Лист1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRanger1 = r1.UsedRange;

                r1.Cells[3, 3] = grika;
                r1.Cells[3, 2] = dk85;
                r1.Cells[3, 5] = rnr;
                r1.Cells[3, 4] = snr;

                sheet2.Close(true, Type.Missing, Type.Missing);
                excel2.Quit();

                //Вставка наладка Rika
                Microsoft.Office.Interop.Excel.Application excel4 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet4 = excel4.Workbooks.Open((Path.Combine(mypath, @"rika\Rika.xlsx")), ReadOnly: false);
                Microsoft.Office.Interop.Excel.Worksheet r4 = excel4.Sheets["Лист1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRanger4 = r4.UsedRange;

                r4.Cells[3, 3] = grika;
                r4.Cells[3, 2] = dk85;
                r4.Cells[3, 5] = rnr;
                r4.Cells[3, 4] = snr;

                sheet4.Close(true, Type.Missing, Type.Missing);
                excel4.Quit();

                //Вставка трафарет Rika30
                Microsoft.Office.Interop.Excel.Application excel3 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet3 = excel3.Workbooks.Open((Path.Combine(mypath, @"rika30\Trafrika30.xlsx")), ReadOnly: false);
                Microsoft.Office.Interop.Excel.Worksheet r3 = excel3.Sheets["Лист1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRanger3 = r3.UsedRange;

                r3.Cells[5, 2] = k625n;
                r3.Cells[5, 4] = snr;
                r3.Cells[5, 6] = rnr;
                
                sheet3.Close(true, Type.Missing, Type.Missing);
                excel3.Quit();

                //Вставка трафарет Rika
                Microsoft.Office.Interop.Excel.Application excel5 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet5 = excel5.Workbooks.Open((Path.Combine(mypath, @"rika\Trafrika.xlsx")), ReadOnly: false);
                Microsoft.Office.Interop.Excel.Worksheet r5 = excel5.Sheets["Лист1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRanger5 = r5.UsedRange;

                r5.Cells[5, 2] = k625n;
                r5.Cells[5, 4] = snr;
                r5.Cells[5, 6] = rnr;

                sheet5.Close(true, Type.Missing, Type.Missing);
                excel5.Quit();

                //Вставка трафарет окончательной шлифовки
                Microsoft.Office.Interop.Excel.Application excel6 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet6 = excel6.Workbooks.Open((Path.Combine(mypath, @"okshlif\okshlif.xlsx")), ReadOnly: false);
                Microsoft.Office.Interop.Excel.Worksheet r6 = excel6.Sheets["Лист1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRanger6 = r6.UsedRange;

                r6.Cells[5, 2] = k625n;
                r6.Cells[5, 4] = sn;
                r6.Cells[5, 6] = rn;

                sheet6.Close(true, Type.Missing, Type.Missing);
                excel6.Quit();

                //Вставка трафарет окончательной шлифовки
                Microsoft.Office.Interop.Excel.Application excel7 = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet7 = excel7.Workbooks.Open((Path.Combine(mypath, @"okshlif30\okshlif30.xlsx")), ReadOnly: false);
                Microsoft.Office.Interop.Excel.Worksheet r7 = excel7.Sheets["Лист1"] as Microsoft.Office.Interop.Excel.Worksheet;

                Excel.Range userRanger7 = r7.UsedRange;

                r7.Cells[5, 2] = k625n;
                r7.Cells[5, 4] = sn;
                r7.Cells[5, 6] = rn;

                sheet7.Close(true, Type.Missing, Type.Missing);
                excel7.Quit();

                label18.Text = "ПУ создан";
                this.label18.ForeColor = System.Drawing.Color.Green;
                button3.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
                button7.Enabled = true;
            }
            catch
            { MessageBox.Show("А в полях точно все цифры?", "Ошибонька"); }
            
        }
        //Блок замены точек запятыми
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            bool text1 = textBox1.Text.Contains(".");
            if (text1)
            {
                textBox1.Text = textBox1.Text.Replace(".", ",");
                textBox1.SelectionStart = textBox1.Text.Length;
            }
        }
        
        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            bool text2 = textBox2.Text.Contains(".");
            if (text2)
            {
                textBox2.Text = textBox2.Text.Replace(".", ",");
                textBox2.SelectionStart = textBox2.Text.Length;
            }
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            bool text3 = textBox3.Text.Contains(".");
            if (text3)
            {
                textBox3.Text = textBox3.Text.Replace(".", ",");
                textBox3.SelectionStart = textBox3.Text.Length;
            }
        }

        private void textBox4_TextChanged_1(object sender, EventArgs e)
        {
            bool text4 = textBox4.Text.Contains(".");
            if (text4)
            {
                textBox4.Text = textBox4.Text.Replace(".", ",");
                textBox4.SelectionStart = textBox4.Text.Length;
            }
        }

        private void textBox5_TextChanged_1(object sender, EventArgs e)
        {
            bool text5 = textBox5.Text.Contains(".");
            if (text5)
            {
                textBox5.Text = textBox5.Text.Replace(".", ",");
                textBox5.SelectionStart = textBox5.Text.Length;
            }
        }

        private void textBox12_TextChanged_1(object sender, EventArgs e)
        {
            bool text12 = textBox12.Text.Contains(".");
            if (text12)
            {
                textBox12.Text = textBox12.Text.Replace(".", ",");
                textBox12.SelectionStart = textBox12.Text.Length;
            }
        }

        private void textBox13_TextChanged_1(object sender, EventArgs e)
        {
            bool text13 = textBox13.Text.Contains(".");
            if (text13)
            {
                textBox13.Text = textBox13.Text.Replace(".", ",");
                textBox13.SelectionStart = textBox13.Text.Length;
            }
        }

        private void textBox6_TextChanged_1(object sender, EventArgs e)
        {
            bool text6 = textBox6.Text.Contains(".");
            if (text6)
            {
                textBox6.Text = textBox6.Text.Replace(".", ",");
                textBox6.SelectionStart = textBox6.Text.Length;
            }
        }

        private void textBox7_TextChanged_1(object sender, EventArgs e)
        {
            bool text7 = textBox7.Text.Contains(".");
            if (text7)
            {
                textBox7.Text = textBox7.Text.Replace(".", ",");
                textBox7.SelectionStart = textBox7.Text.Length;
            }
        }

        private void textBox8_TextChanged_1(object sender, EventArgs e)
        {
            bool text8 = textBox1.Text.Contains(".");
            if (text8)
            {
                textBox8.Text = textBox8.Text.Replace(".", ",");
                textBox8.SelectionStart = textBox8.Text.Length;
            }
        }

        private void textBox9_TextChanged_1(object sender, EventArgs e)
        {
            bool text9 = textBox9.Text.Contains(".");
            if (text9)
            {
                textBox9.Text = textBox9.Text.Replace(".", ",");
                textBox9.SelectionStart = textBox9.Text.Length;
            }
        }

        private void textBox10_TextChanged_1(object sender, EventArgs e)
        {
            bool text10 = textBox10.Text.Contains(".");
            if (text10)
            {
                textBox10.Text = textBox10.Text.Replace(".", ",");
                textBox10.SelectionStart = textBox10.Text.Length;
            }
        }

        private void textBox11_TextChanged_1(object sender, EventArgs e)
        {
            bool text11 = textBox11.Text.Contains(".");
            if (text11)
            {
                textBox11.Text = textBox11.Text.Replace(".", ",");
                textBox11.SelectionStart = textBox11.Text.Length;
            }
        }
        //
        //

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)//Кнопка выход
        {
            this.Hide();
        }

        //Блок печати
        private void button3_Click(object sender, EventArgs e)
        {
            string nas = textBox14.Text;
            string ind = textBox15.Text;
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Path.Combine(mypath, @"pu\" + nas + ind + " " + dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy") + ".xlsx"), ReadOnly: false);


                if (radioButton1.Checked == true)//9
                {
                    if (radioButton6.Checked == true)//дорн
                    {
                        excel.ActiveWorkbook.PrintOutEx(6, 7);
                        excel.ActiveWorkbook.PrintOutEx(22, 22);
                        excel.ActiveWorkbook.PrintOutEx(29, 31);
                        excel.ActiveWorkbook.PrintOutEx(34, 35);
                    }
                    if (radioButton5.Checked == true)//без дорна
                    {
                        excel.ActiveWorkbook.PrintOutEx(16, 17);
                        excel.ActiveWorkbook.PrintOutEx(22, 22);
                        excel.ActiveWorkbook.PrintOutEx(30, 31);
                        excel.ActiveWorkbook.PrintOutEx(34, 35);
                    }                    
                 }

                if (radioButton2.Checked == true)//7
                {
                    if (radioButton6.Checked == true)//дорн
                    {
                        excel.ActiveWorkbook.PrintOutEx(1, 2);
                        excel.ActiveWorkbook.PrintOutEx(22, 22);
                        excel.ActiveWorkbook.PrintOutEx(29, 32);
                        excel.ActiveWorkbook.PrintOutEx(34, 35);
                    }
                    if (radioButton5.Checked == true)//без дорна
                    {
                        excel.ActiveWorkbook.PrintOutEx(11, 12);
                        excel.ActiveWorkbook.PrintOutEx(22, 22);
                        excel.ActiveWorkbook.PrintOutEx(30, 32);
                        excel.ActiveWorkbook.PrintOutEx(34, 35);
                    }
                }

                if (radioButton8.Checked == true)//30гр
                 {
                    excel.ActiveWorkbook.PrintOutEx(23, 23);
                    excel.ActiveWorkbook.PrintOutEx(46, 46);
                }
                 if (radioButton7.Checked == true)//без 30гр
                {
                     excel.ActiveWorkbook.PrintOutEx(24, 24);
                     excel.ActiveWorkbook.PrintOutEx(47, 47);
                }

                if (radioButton10.Checked == true)//2 фаски
                {
                    excel.ActiveWorkbook.PrintOutEx(33, 33);
                }

                if (radioButton4.Checked == true)//большая выточка
                {
                    if (radioButton6.Checked == true)//дорн
                    {
                        excel.ActiveWorkbook.PrintOutEx(25, 25);
                    }
                    if (radioButton5.Checked == true)//без дорна
                    {
                        excel.ActiveWorkbook.PrintOutEx(26, 26);
                    }
                }

                if (radioButton3.Checked == true)//маленькая выточка
                {
                    if (radioButton6.Checked == true)//дорн
                    {
                        excel.ActiveWorkbook.PrintOutEx(27, 27);
                    }
                    if (radioButton5.Checked == true)//без дорна
                    {
                        excel.ActiveWorkbook.PrintOutEx(26, 28);
                    }
                }

                if (radioButton2.Checked == true)//7
                {
                    if (radioButton4.Checked == true)//большая выточка
                    {
                        excel.ActiveWorkbook.PrintOutEx(36, 36);
                    }
                    if (radioButton3.Checked == true)//маленькая выточка
                    {
                        excel.ActiveWorkbook.PrintOutEx(38, 38);
                    }
                }

                if (radioButton1.Checked == true)//9
                {
                    if (radioButton4.Checked == true)//большая выточка
                    {
                        excel.ActiveWorkbook.PrintOutEx(37, 37);
                    }
                    if (radioButton3.Checked == true)//маленькая выточка
                    {
                        excel.ActiveWorkbook.PrintOutEx(39, 39);
                    }
                }

                if (radioButton2.Checked == true)//7
                {
                    if (radioButton11.Checked == true)//без супфины
                    {
                        excel.ActiveWorkbook.PrintOutEx(40, 41);
                        excel.ActiveWorkbook.PrintOutEx(45, 45);
                    }
                    if (radioButton12.Checked == true)//супфина
                    {
                        excel.ActiveWorkbook.PrintOutEx(43, 43);
                        excel.ActiveWorkbook.PrintOutEx(45, 45);
                    }
                }

                if (radioButton1.Checked == true)//9
                {
                    if (radioButton11.Checked == true)//без супфины
                    {
                        excel.ActiveWorkbook.PrintOutEx(40, 40);
                        excel.ActiveWorkbook.PrintOutEx(42, 42);
                    }
                    if (radioButton12.Checked == true)//супфина
                    {
                        excel.ActiveWorkbook.PrintOutEx(44, 44);
                    }
                }
                MessageBox.Show("Голоса в моей голове и логика программы приказали мне перепутать все листы при выводе на печать. Страдай.", "Как-то так");
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
            }
            catch { MessageBox.Show("Что-то пошло не так.", "Ошибонька"); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start(Path.Combine(mypath, @"pu\"));
        }

        private void операцииПоДругомуПУToolStripMenuItem_Click(object sender, EventArgs e)
        {
           Form2 fr2 = new Form2();
            fr2.textBox1.Text = textBox14.Text;
            fr2.textBox2.Text = textBox15.Text;
            fr2.Show();
        }
        
        //Кнопки открытия солида
        private void button4_Click(object sender, EventArgs e)//Открыть карту наладки Rika
        {
            if (textBox1.Text != "0" & textBox2.Text != "0" & textBox3.Text != "0" & textBox7.Text != "0" & textBox8.Text != "0" & textBox6.Text != "0")
            {
                MessageBox.Show("Не забудь поменять название распылителя в карте наладки!", "Как-то так");
                if (radioButton7.Checked == true)//Сфера без 30гр
                { Process.Start(Path.Combine(mypath, @"rika\Rika.SLDDRW")); }

                if (radioButton8.Checked == true)//Сфера 30гр
                { Process.Start(Path.Combine(mypath, @"rika30\Rika30.SLDDRW")); }
            }
            else { MessageBox.Show("Не хватает данных для карты наладки. Заполни поля и пересоздай ПУ.", "Как-то так"); }
        }

        private void button5_Click(object sender, EventArgs e)//Открыть трафарет Rika
        {
            if (textBox7.Text != "0" & textBox8.Text != "0" & textBox6.Text != "0")
            {
                MessageBox.Show("Не забудь поменять название распылителя в трафарете!", "Как-то так");
                if (radioButton7.Checked == true)//Сфера без 30гр
                { Process.Start(Path.Combine(mypath, @"rika\trafrika.SLDDRW")); }

                if (radioButton8.Checked == true)//Сфера 30гр
                { Process.Start(Path.Combine(mypath, @"rika30\trafrika30.SLDDRW")); }
            }
            else { MessageBox.Show("Не хватает данных для трафарета. Заполни поля и пересоздай ПУ.", "Как-то так"); }
        }

        private void button7_Click(object sender, EventArgs e)//Открыть трафарет окончательной шлифовки
        {
            if (textBox7.Text != "0" & textBox8.Text != "0" & textBox6.Text != "0")
            {
                MessageBox.Show("Не забудь поменять название распылителя в трафарете!", "Как-то так");
                if (radioButton7.Checked == true)//Сфера без 30гр
                { Process.Start(Path.Combine(mypath, @"okshlif\okshlif.SLDDRW")); }

                if (radioButton8.Checked == true)//Сфера 30гр
                { Process.Start(Path.Combine(mypath, @"okshlif30\okshlif30.SLDDRW")); }
            }
            else { MessageBox.Show("Не хватает данных для трафарета. Заполни поля и пересоздай ПУ.", "Как-то так"); }
        }

       
    }
}
