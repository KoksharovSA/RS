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


namespace RS
{
    public partial class Form1 : Form
    {
        private OleDbConnection con = new OleDbConnection();
        private OleDbConnection con1 = new OleDbConnection();
        public Form1()
        {
            InitializeComponent();
        }
        public string mypath;
        
        private void Form1_Load(object sender, EventArgs e)
        {
            mypath = Directory.GetCurrentDirectory();                     
            con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Ts-03\\users\\OGK\\Кокшаров С.А\\db\\DB.accdb;Jet OLEDB:Database Password=6567604";
            con.Open();

            OleDbCommand cmd4 = new OleDbCommand();
            cmd4.Connection = con;
            cmd4.CommandText = "insert into stat (dat, kto) values ('"+ Convert.ToString(DateTime.Now) + "','" + label25.Text + "')";
            cmd4.ExecuteNonQuery();                      

            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = con;
            string query = "SELECT fio1 from komu";
            cmd.CommandText = query;
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            { comboBox1.Items.Add(reader[0].ToString());}

            OleDbCommand cmd2 = new OleDbCommand();
            cmd2.Connection = con;
            string query2 = "SELECT fio2 from ot";
            cmd2.CommandText = query2;
            OleDbDataReader reader2 = cmd2.ExecuteReader();
            while (reader2.Read())
            {comboBox2.Items.Add(reader2[0].ToString());}

            OleDbCommand cmd3 = new OleDbCommand();
            cmd3.Connection = con;
            string query3 = "SELECT ispol from isp";
            cmd3.CommandText = query3;
            OleDbDataReader reader3 = cmd3.ExecuteReader();
            while (reader3.Read())
            { comboBox3.Items.Add(reader3[0].ToString()); }

            
            axAcroPDF1.LoadFile(Path.Combine(mypath, "pas.pdf"));

            //ПУ

            OleDbCommand cmdsog = new OleDbCommand();
            cmdsog.Connection = con;
            string querysog = "SELECT sog from soglas";
            cmdsog.CommandText = querysog;
            OleDbDataReader readersog = cmdsog.ExecuteReader();
            while (readersog.Read())
            { comboBox22.Items.Add(readersog[0].ToString()); }

            OleDbCommand cmdra = new OleDbCommand();
            cmdra.Connection = con;
            string queryra = "SELECT ra from sher";
            cmdra.CommandText = queryra;
            OleDbDataReader readerra = cmdra.ExecuteReader();
            while (readerra.Read())
            { comboBox23.Items.Add(readerra[0].ToString()); }

            OleDbCommand cmdmat = new OleDbCommand();
            cmdmat.Connection = con;
            string querymat = "SELECT stil from met";
            cmdmat.CommandText = querymat;
            OleDbDataReader readermat = cmdmat.ExecuteReader();
            while (readermat.Read())
            { comboBox4.Items.Add(readermat[0].ToString()); }

            OleDbCommand cmdceh = new OleDbCommand();
            cmdceh.Connection = con;
            string queryceh = "SELECT ceh from ceh";
            cmdceh.CommandText = queryceh;
            OleDbDataReader readerceh = cmdceh.ExecuteReader();
            while (readerceh.Read())
            { comboBox5.Items.Add(readerceh[0].ToString()); }

            OleDbCommand cmdoper = new OleDbCommand();
            cmdoper.Connection = con;
            string queryoper = "SELECT oper from oper";
            cmdoper.CommandText = queryoper;
            OleDbDataReader readeroper = cmdoper.ExecuteReader();
            while (readeroper.Read())
            { comboBox7.Items.Add(readeroper[0].ToString()); }

            OleDbCommand cmdobor = new OleDbCommand();
            cmdobor.Connection = con;
            string queryobor = "SELECT obor from obor";
            cmdobor.CommandText = queryobor;
            OleDbDataReader readerobor = cmdobor.ExecuteReader();
            while (readerobor.Read())
            { comboBox8.Items.Add(readerobor[0].ToString()); }

            OleDbCommand cmdprov = new OleDbCommand();
            cmdprov.Connection = con;
            string queryprov = "SELECT prov from prov";
            cmdprov.CommandText = queryprov;
            OleDbDataReader readerprov = cmdprov.ExecuteReader();
            while (readerprov.Read())
            { comboBox9.Items.Add(readerprov[0].ToString()); }

            OleDbCommand cmdrazr = new OleDbCommand();
            cmdrazr.Connection = con;
            string queryrazr = "SELECT razr from razr";
            cmdrazr.CommandText = queryrazr;
            OleDbDataReader readerrazr = cmdrazr.ExecuteReader();
            while (readerrazr.Read())
            { comboBox6.Items.Add(readerrazr[0].ToString()); }

            OleDbCommand cmdrinst = new OleDbCommand();
            cmdrinst.Connection = con;
            string queryrinst = "SELECT rinst from rinst";
            cmdrinst.CommandText = queryrinst;
            OleDbDataReader readerrinst = cmdrinst.ExecuteReader();
            while (readerrinst.Read())
            {
                comboBox10.Items.Add(readerrinst[0].ToString());
                comboBox11.Items.Add(readerrinst[0].ToString());
                comboBox12.Items.Add(readerrinst[0].ToString());
                comboBox13.Items.Add(readerrinst[0].ToString());
                comboBox14.Items.Add(readerrinst[0].ToString());
                comboBox15.Items.Add(readerrinst[0].ToString());
            }

            OleDbCommand cmdvinst = new OleDbCommand();
            cmdvinst.Connection = con;
            string queryvinst = "SELECT vinst from vinst";
            cmdvinst.CommandText = queryvinst;
            OleDbDataReader readervinst = cmdvinst.ExecuteReader();
            while (readervinst.Read())
            {
                comboBox16.Items.Add(readervinst[0].ToString());
                comboBox17.Items.Add(readervinst[0].ToString());
                comboBox18.Items.Add(readervinst[0].ToString());
                comboBox19.Items.Add(readervinst[0].ToString());
                comboBox20.Items.Add(readervinst[0].ToString());
                comboBox21.Items.Add(readervinst[0].ToString());
            }

            OleDbCommand cmdiinst = new OleDbCommand();
            cmdiinst.Connection = con;
            string queryiinst = "SELECT iinst from iinst";
            cmdiinst.CommandText = queryiinst;
            OleDbDataReader readeriinst = cmdiinst.ExecuteReader();
            while (readeriinst.Read())
            {
                comboBox24.Items.Add(readeriinst[0].ToString());
                comboBox25.Items.Add(readeriinst[0].ToString());
                comboBox26.Items.Add(readeriinst[0].ToString());
                comboBox27.Items.Add(readeriinst[0].ToString());
            }

            con.Close();

        } //Загрузка формы

        private void button1_Click(object sender, EventArgs e)
        {
            string komu = label6.Text;
            string fio1 = comboBox1.Text;
            string ot = label7.Text;
            string fio2 = comboBox2.Text;
            string dat = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
            string ispol = comboBox3.Text;
            string tekst = "     "+richTextBox1.Text;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open((Path.Combine(mypath, "SZ.xlsx")), ReadOnly: false);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["variable"] as Microsoft.Office.Interop.Excel.Worksheet;

            Excel.Range userRange = x.UsedRange;
            
            x.Cells[2, 1] = komu;
            x.Cells[2, 2] = fio1;
            x.Cells[2, 3] = dat;
            x.Cells[2, 4] = tekst;
            x.Cells[2, 5] = ot;
            x.Cells[2, 6] = fio2;
            x.Cells[2, 7] = ispol;
                       
            excel.ActiveWorkbook.PrintOutEx(1, 1);
                        
            sheet.Close(true, (Path.Combine(mypath, @"sz\"+fio1 + " "+dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy HH'.'mm") + ".xlsx")), Type.Missing);
            excel.Quit();

        } //Печатать служебную

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd1 = new OleDbCommand();
            cmd1.Connection = con;
            string query1 = "SELECT * from komu where fio1='" + comboBox1.Text + "'";
            cmd1.CommandText = query1;
            OleDbDataReader reader1 = cmd1.ExecuteReader();
            while (reader1.Read())
            {
                label6.Text = reader1["komu"].ToString();
            }

            
            con.Close();
        } 

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd2 = new OleDbCommand();
            cmd2.Connection = con;
            string query2 = "SELECT * from ot where fio2='" + comboBox2.Text + "'";
            cmd2.CommandText = query2;
            OleDbDataReader reader2 = cmd2.ExecuteReader();
            while (reader2.Read())
            {
                label7.Text = reader2["ot"].ToString();
            }
            con.Close();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            con.Close();
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start(@"\\Ts-03\users\OGK\Реестр Заданий\РЕЕСТР ЗАДАНИЙ.xlsx");
        } //Реестр заданий

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {            
            con.Open();
            OleDbCommand cmdiot = new OleDbCommand();
            cmdiot.Connection = con;
            string queryiot = "SELECT * from oper where oper='" + comboBox7.Text + "'";
            cmdiot.CommandText = queryiot;
            OleDbDataReader readeriot = cmdiot.ExecuteReader();
            while (readeriot.Read())
            { textBox8.Text = readeriot["iot"].ToString(); }

            OleDbCommand cmdtekstoper = new OleDbCommand();
            cmdtekstoper.Connection = con;
            string querytekstoper = "SELECT * from oper where oper='" + comboBox7.Text + "'";
            cmdtekstoper.CommandText = querytekstoper;
            OleDbDataReader readertekstoper = cmdtekstoper.ExecuteReader();
            while (readertekstoper.Read())
            { textBox6.Text = readertekstoper["tekstoper"].ToString(); }
                       
            con.Close();
            checkBox5.Checked = true;
            checkBox6.Checked = false;
            if (comboBox7.Text == "Термообработка") { checkBox5.Checked = false; checkBox6.Checked = true; }
            if (comboBox7.Text == "Моечная") { checkBox5.Checked = false; checkBox6.Checked = false; }
            if (comboBox7.Text == "Транспортная") { checkBox5.Checked = false; checkBox6.Checked = false; }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                //con.Open();
                //OleDbCommand cmd = new OleDbCommand();
                //cmd.Connection = con;
                //cmd.CommandText = "insert into sl (nomer, inddet, nasdet, metal, proekt, zadan, data, kto) values ('" + textBox11.Text + "','" + textBox4.Text + "','" + textBox3.Text + "','" + comboBox4.Text + "','" + textBox9.Text + "','" + textBox10.Text + "','" + dateTimePicker2.Text + "','" + label25.Text + "')";
                //cmd.ExecuteNonQuery();
                //con.Close();

                string nas = textBox3.Text;
                string ind = textBox4.Text;

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open((Path.Combine(mypath, "mk.xlsm")), ReadOnly: false, Password: "6567604");
                                               
                sheet.Close(true, (Path.Combine(mypath, @"mk\" + "PU" + ind + nas + ".xlsm")), Type.Missing);
                excel.Quit();

                //File.Copy(Path.Combine(mypath, "mk.xlsx"), Path.Combine(mypath, @"mk\" + "PU" + ind + nas + ".xlsx"), true);
                label23.Text = "Файл плана управления создан";
                this.label23.ForeColor = System.Drawing.Color.Green;
            }
            catch (Exception){ MessageBox.Show("Что-то пошло не так. Может файл уже открыт?", "Ошибонька"); }
        } //Создать МК

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                button5.Enabled = false;
                progressBar1.Value = 0;
                string ra = comboBox23.Text;
                string sog = comboBox22.Text;
                string proekt = textBox9.Text;
                string zad = textBox10.Text;
                string listov = textBox1.Text;
                string mat = comboBox4.Text;
                string inv = textBox2.Text;
                string nas = textBox3.Text;
                string ind = textBox4.Text;
                string noper = textBox5.Text;
                string noperpos = Convert.ToString(numericUpDown1.Value);
                string noperpos2 = Convert.ToString(numericUpDown2.Value);
                string noperpos3 = Convert.ToString(numericUpDown3.Value);
                string ceh = comboBox5.Text;
                string oper = comboBox7.Text;
                string iot = textBox8.Text;
                string obor = comboBox8.Text;
                string tekstoper = textBox6.Text;
                string prov = comboBox9.Text;
                string razr = comboBox6.Text;
                string dat = dateTimePicker2.Value.ToString("dd'.'MM'.'yyyy");

                string ri1 = comboBox10.Text;
                string ri2 = comboBox11.Text;
                string ri3 = comboBox12.Text;
                string ri4 = comboBox13.Text;
                string ri5 = comboBox14.Text;
                string ri6 = comboBox15.Text;

                string vi1 = comboBox21.Text;
                string vi2 = comboBox20.Text;
                string vi3 = comboBox19.Text;
                string vi4 = comboBox18.Text;
                string vi5 = comboBox17.Text;
                string vi6 = comboBox16.Text;

                string ii1 = comboBox27.Text;
                string ii2 = comboBox26.Text;
                string ii3 = comboBox25.Text;
                string ii4 = comboBox24.Text;
                progressBar1.Value = 20;

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open((Path.Combine(mypath, @"mk\" + "PU" + ind + nas + ".xlsm")), ReadOnly: false);
                if (checkBox4.Checked == true)
                {
                    //Вставка значений в МК и СЛ
                    int a = 5;
                    int b = 6;
                    int c = 7;
                    int s = 13;
                    if (Convert.ToInt32(noperpos) == 1) { s = 13; }
                    else
                    {
                        if (Convert.ToInt32(noperpos) <= 30) { for (int r = 2; r <= 30; r++) { if (Convert.ToInt32(noperpos) == r) { s = 13 + ((2 * (Convert.ToInt32(noperpos)) - 2)); break; } } }
                        else
                        {
                            if (Convert.ToInt32(noperpos) <= 55) { for (int r = 31; r <= 55; r++) { if (Convert.ToInt32(noperpos) == r) { s = 13 + ((2 * (Convert.ToInt32(noperpos)) + 1)); break; } } }
                        }
                    }
                    if (Convert.ToInt32(noperpos) == 1) { a = 5; b = a + 1; c = b + 1; }
                    else
                    {
                        if (Convert.ToInt32(noperpos) <= 13) { for (int r = 2; r <= 13; r++) { if (Convert.ToInt32(noperpos) == r) { a = a + (4 * (Convert.ToInt32(noperpos) - 1)); b = a + 1; c = b + 1; break; } } }
                        else
                        {
                            if (Convert.ToInt32(noperpos) <= 26) { for (int r = 14; r <= 26; r++) { if (Convert.ToInt32(noperpos) == r) { a = a + (4 * (Convert.ToInt32(noperpos) - 1)) + 7; b = a + 1; c = b + 1; break; } } }
                            else
                            {
                                if (Convert.ToInt32(noperpos) <= 39) { for (int r = 27; r <= 39; r++) { if (Convert.ToInt32(noperpos) == r) { a = a + (4 * (Convert.ToInt32(noperpos) - 1)) + 14; b = a + 1; c = b + 1; break; } } }
                                else
                                {
                                    if (Convert.ToInt32(noperpos) <= 52) { for (int r = 40; r <= 52; r++) { if (Convert.ToInt32(noperpos) == r) { a = a + (4 * (Convert.ToInt32(noperpos) - 1)) + 21; b = a + 1; c = b + 1; break; } } }
                                    else
                                    {
                                        if (Convert.ToInt32(noperpos) <= 65) { for (int r = 53; r <= 65; r++) { if (Convert.ToInt32(noperpos) == r) { a = a + (4 * (Convert.ToInt32(noperpos) - 1)) + 28; b = a + 1; c = b + 1; break; } } }
                                    }
                                }
                            }
                        }
                    }
                    progressBar1.Value = 40;

                    Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["mk"] as Microsoft.Office.Interop.Excel.Worksheet;

                    Excel.Range userRange = x.UsedRange;

                    x.Cells[59, 7] = sog;
                    x.Cells[2, 4] = listov;
                    x.Cells[2, 5] = mat;
                    x.Cells[1, 7] = nas;
                    x.Cells[2, 7] = ind;
                    x.Cells[2, 9] = inv;
                    x.Cells[57, 7] = razr;
                    x.Cells[58, 7] = prov;
                    x.Cells[57, 9] = dat;
                    x.Cells[58, 9] = dat;
                    x.Cells[57, 1] = nas;
                    x.Cells[58, 1] = ind;
                    x.Cells[a, 1] = noper;
                    x.Cells[a, 2] = oper;
                    x.Cells[a, 7] = iot;
                    x.Cells[a, 9] = ceh;
                    x.Cells[b, 2] = obor;
                    x.Cells[c, 2] = tekstoper;

                    if (checkBox8.Checked == true)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet u = excel.Sheets["sl"] as Microsoft.Office.Interop.Excel.Worksheet;

                        Excel.Range userRangeu = u.UsedRange;

                        u.Cells[1, 7] = zad;
                        u.Cells[2, 7] = proekt;
                        u.Cells[7, 2] = mat;
                        u.Cells[4, 6] = nas;
                        u.Cells[4, 1] = ind;
                        u.Cells[s, 1] = noper;
                        u.Cells[s, 2] = oper;
                        u.Cells[s, 5] = ceh;
                    }

                    progressBar1.Value = 50;
                    numericUpDown1.Value = Convert.ToInt32(noperpos) + 1;
                }

                if (checkBox5.Checked == true)
                {
                    //Вставка значений в ПУ
                    int d = 0;
                    if (Convert.ToInt32(noperpos2) == 1) { d = 0; }
                    else
                    {
                        for (int u = 2; u <= 60; u++)
                        {
                            if (Convert.ToInt32(noperpos2) == u) { d = d + (60 * (Convert.ToInt32(noperpos2) - 1)); break; }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet y = excel.Sheets["pu"] as Microsoft.Office.Interop.Excel.Worksheet;

                    Excel.Range userRange1 = y.UsedRange;

                    y.Cells[5 + d, 15] = ra;
                    y.Cells[60 + d, 11] = sog;
                    y.Cells[2 + d, 2] = mat;
                    y.Cells[1 + d, 15] = inv;
                    y.Cells[58 + d, 11] = razr;
                    y.Cells[59 + d, 11] = prov;
                    y.Cells[58 + d, 15] = dat;
                    y.Cells[59 + d, 15] = dat;
                    y.Cells[60 + d, 15] = dat;
                    y.Cells[58 + d, 1] = nas;
                    y.Cells[59 + d, 1] = ind;
                    y.Cells[3 + d, 1] = noper;
                    y.Cells[1 + d, 2] = oper;
                    y.Cells[1 + d, 13] = iot;
                    y.Cells[3 + d, 2] = obor;
                    y.Cells[49 + d, 1] = ri1;
                    y.Cells[50 + d, 1] = ri2;
                    y.Cells[51 + d, 1] = ri3;
                    y.Cells[52 + d, 1] = ri4;
                    y.Cells[53 + d, 1] = ri5;
                    y.Cells[54 + d, 1] = ri6;
                    y.Cells[49 + d, 5] = vi1;
                    y.Cells[50 + d, 5] = vi2;
                    y.Cells[51 + d, 5] = vi3;
                    y.Cells[52 + d, 5] = vi4;
                    y.Cells[53 + d, 5] = vi5;
                    y.Cells[54 + d, 5] = vi6;
                    y.Cells[49 + d, 11] = ii1;
                    y.Cells[51 + d, 11] = ii2;
                    y.Cells[53 + d, 11] = ii3;
                    y.Cells[55 + d, 11] = ii4;

                    //excel.Run("mak");
                    
                    numericUpDown1.Value = Convert.ToInt32(noperpos) + 1;
                    numericUpDown2.Value = Convert.ToInt32(noperpos2) + 1;
                }

                if (checkBox6.Checked == true)
                {
                    //Вставка значений в ТО

                    int d = 0;
                    if (Convert.ToInt32(noperpos3) == 1) { d = 0; }
                    else
                    {
                        for (int u = 2; u <= 5; u++)
                        {
                            if (Convert.ToInt32(noperpos3) == u) { d = d + (60 * (Convert.ToInt32(noperpos3) - 1)); break; }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Worksheet y = excel.Sheets["to"] as Microsoft.Office.Interop.Excel.Worksheet;

                    Excel.Range userRange1 = y.UsedRange;

                    y.Cells[3 + d, 1] = noper;
                    numericUpDown3.Value = Convert.ToInt32(noperpos3) + 1;
                    numericUpDown1.Value = Convert.ToInt32(noperpos) + 1;
                }

                    Microsoft.Office.Interop.Excel.Worksheet z = excel.Sheets["tl"] as Microsoft.Office.Interop.Excel.Worksheet;

                    Excel.Range userRange2 = z.UsedRange;

                    z.Cells[1, 13] = inv;
                    z.Cells[3, 13] = zad;
                    z.Cells[4, 13] = proekt;
                    z.Cells[20, 6] = nas;
                    z.Cells[18, 6] = ind;
                    z.Cells[23, 13] = mat;
                    progressBar1.Value = 90;
                    sheet.Close(true, Type.Missing, Type.Missing);
                    excel.Quit();
                                
                progressBar1.Value=100;
                numericUpDown1.Value = Convert.ToInt32(noperpos) + 1;
                button5.Enabled = true;
                if (textBox7.Text == "") { textBox7.Text = textBox5.Text + " " + comboBox7.Text; }
                else { textBox7.Text = textBox7.Text + Environment.NewLine + textBox5.Text + " " + comboBox7.Text; }
                textBox7.SelectionStart = textBox7.Text.Length;
                textBox7.ScrollToCaret();
                if (Convert.ToInt32(textBox5.Text) < 99)
                { textBox5.Text = "0" + Convert.ToString(Convert.ToInt32(textBox5.Text) + 5); }
                else
                { textBox5.Text = Convert.ToString(Convert.ToInt32(textBox5.Text) + 5); }
            }
            catch (Exception){ MessageBox.Show("Что-то пошло не так. Может не создал файл.", "Ошибонька"); }

        } //Добавить в ПУ

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                string nas = textBox3.Text;
                string ind = textBox4.Text;
                string listov = textBox1.Text;
                string noperpos = Convert.ToString(numericUpDown1.Value);
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open((Path.Combine(mypath, @"mk\" + "PU" + ind + nas + ".xlsx")), ReadOnly: false);

                if (checkBox2.Checked == true)
                {
                    Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["mk"] as Microsoft.Office.Interop.Excel.Worksheet;
                    excel.ActiveWorkbook.PrintOutEx(1, listov);
                }
                if (checkBox3.Checked == true)
                {
                    if(MessageBox.Show("Количество карт эскизов на печать " + noperpos + ". Продолжить печать? (если учесть что она нахуй никому не нужна без картинок)", "Печать", MessageBoxButtons.YesNo)==DialogResult.Yes)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["pu"] as Microsoft.Office.Interop.Excel.Worksheet;
                        excel.ActiveWorkbook.PrintOutEx(6, 5 + noperpos);
                    }
                }
                if (checkBox1.Checked == true)
                {
                    Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["tl"] as Microsoft.Office.Interop.Excel.Worksheet;
                    excel.ActiveWorkbook.PrintOutEx(30, 30);
                }
                if (checkBox7.Checked == true)
                {
                    Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["sl"] as Microsoft.Office.Interop.Excel.Worksheet;
                    excel.ActiveWorkbook.PrintOutEx(62, 63);
                }
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
            }
            catch (Exception) { MessageBox.Show("Что-то пошло не так. Может не создал файл.", "Ошибонька"); }
            
        } //Печать ПУ

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start(@"\\Ts-03\users\OGK\Чудинов А. С\Bas\PU.xlsx");
        } //Открыть реестр ПУ

        private void pURToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                if (label25.Text == "admin")
                {
                    PUR.Form2 f2 = new PUR.Form2();
                    f2.ShowDialog();
                }
                con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Ts-03\\users\\OGK\\Кокшаров С.А\\db\\DB.accdb;Jet OLEDB:Database Password=6567604";
                con.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                string query = "SELECT namelogin, doppur from login";
                cmd.CommandText = query;
                OleDbDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    if (label25.Text == Convert.ToString(reader[0]) & Convert.ToBoolean(reader[1])==true)
                    {
                        PUR.Form2 f2 = new PUR.Form2();
                        f2.ShowDialog();
                        break; 
                    }                   
                }
                con.Close();                              
            }
            catch (Exception) { MessageBox.Show("Что-то произошло", "Ошибка"); }
        } //Открыть PUR

        private void qSONToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(Path.Combine(mypath, "qson.exe"));
            }
            catch (Exception) { MessageBox.Show("Что-то пошло не так. Не могу найти программу.", "Ошибонька"); }
} //Открыть QSON

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox6.Text == "Пузиков") { MessageBox.Show("Артём ты бы трахнул овцу?", "Вопрос на миллион"); }
        }
                
        private void button9_Click(object sender, EventArgs e)
        {
            string komu = label6.Text;
            string fio1 = comboBox1.Text;
            string ot = label7.Text;
            string fio2 = comboBox2.Text;
            string dat = dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy");
            string ispol = comboBox3.Text;
            string tekst = "     " + richTextBox1.Text;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open((Path.Combine(mypath, "SZ.xlsx")), ReadOnly: false);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.Sheets["variable"] as Microsoft.Office.Interop.Excel.Worksheet;

            Excel.Range userRange = x.UsedRange;

            x.Cells[2, 1] = komu;
            x.Cells[2, 2] = fio1;
            x.Cells[2, 3] = dat;
            x.Cells[2, 4] = tekst;
            x.Cells[2, 5] = ot;
            x.Cells[2, 6] = fio2;
            x.Cells[2, 7] = ispol;

            sheet.Close(true, (Path.Combine(mypath, @"sz\" + fio1 + " " + dateTimePicker1.Value.ToString("dd'.'MM'.'yyyy HH'.'mm") + ".xlsx")), Type.Missing);
            excel.Quit();
        }

        private void пУToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox7.Text = "";
            label23.Text = "Файл плана управления не создан";
            this.label23.ForeColor = System.Drawing.Color.Red;
            progressBar1.Value = 0;
            textBox9.Clear();
            textBox10.Clear();
            textBox1.Clear();
            comboBox4.Text="";
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            numericUpDown1.Value=1;
            comboBox5.Text = "";
            comboBox7.Text = "";
            textBox8.Clear();
            comboBox8.Text = "";
            textBox6.Clear();
            comboBox9.Text = "";
            comboBox6.Text = "";
                        
            comboBox11.Text = "";
            comboBox12.Text = "";
            comboBox13.Text = "";
            comboBox14.Text = "";
            comboBox15.Text = "";

            comboBox21.Text = "";
            comboBox20.Text = "";
            comboBox19.Text = "";
            comboBox18.Text = "";
            comboBox17.Text = "";
            comboBox16.Text = "";

            comboBox27.Text = ""; 
            comboBox26.Text = "";
            comboBox25.Text = "";
            comboBox24.Text = "";
        }

        private void служебнаяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label6.Text = "";
            comboBox1.Text = "";
            label7.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            richTextBox1.Text = "";
            
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("   Программа "+this.Text+" Выжимка из лени и хотения чтобы за меня всю работу делал компуктер:)\n    На данный момент она умеет делать планы управления (ТЛ, МК, КЭ, ТО, СЛ) без картинок, печатать служебные и паспорта.\n   Надеюсь будет полезной. \n©Кокшаров Сергей Александрович");
        }

        private void pUR2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (label25.Text == "admin")
                {
                    PUR2.Form1 f2 = new PUR2.Form1();
                    f2.ShowDialog();                   
                }

                con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Ts-03\\users\\OGK\\Кокшаров С.А\\db\\DB.accdb;Jet OLEDB:Database Password=6567604";
                con.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                string query = "SELECT namelogin, doppur2 from login";
                cmd.CommandText = query;
                OleDbDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    if (label25.Text == Convert.ToString(reader[0]) & Convert.ToBoolean(reader[1]) == true)
                    {
                        PUR2.Form1 f2 = new PUR2.Form1();
                        f2.ShowDialog();
                        break;
                    }
                }
                con.Close();
            }
            catch (Exception) { MessageBox.Show("Что-то произошло", "Ошибка"); }
        }

        private void локальнаяБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Если незнаешь что это за кнопка нехуй тыкать. Ты знаешь что это за кнопка?", "Открыть ЛБД", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {

                con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(mypath, "DB.accdb") + ";Jet OLEDB:Database Password=6567604";

                con.Open();

                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                string query = "SELECT fio1 from komu";
                cmd.CommandText = query;
                OleDbDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                { comboBox1.Items.Add(reader[0].ToString()); }

                OleDbCommand cmd2 = new OleDbCommand();
                cmd2.Connection = con;
                string query2 = "SELECT fio2 from ot";
                cmd2.CommandText = query2;
                OleDbDataReader reader2 = cmd2.ExecuteReader();
                while (reader2.Read())
                { comboBox2.Items.Add(reader2[0].ToString()); }

                OleDbCommand cmd3 = new OleDbCommand();
                cmd3.Connection = con;
                string query3 = "SELECT ispol from isp";
                cmd3.CommandText = query3;
                OleDbDataReader reader3 = cmd3.ExecuteReader();
                while (reader3.Read())
                { comboBox3.Items.Add(reader3[0].ToString()); }


                axAcroPDF1.LoadFile(Path.Combine(mypath, "pas.pdf"));

                //ПУ

                OleDbCommand cmdsog = new OleDbCommand();
                cmdsog.Connection = con;
                string querysog = "SELECT sog from soglas";
                cmdsog.CommandText = querysog;
                OleDbDataReader readersog = cmdsog.ExecuteReader();
                while (readersog.Read())
                { comboBox22.Items.Add(readersog[0].ToString()); }

                OleDbCommand cmdra = new OleDbCommand();
                cmdra.Connection = con;
                string queryra = "SELECT ra from sher";
                cmdra.CommandText = queryra;
                OleDbDataReader readerra = cmdra.ExecuteReader();
                while (readerra.Read())
                { comboBox23.Items.Add(readerra[0].ToString()); }

                OleDbCommand cmdmat = new OleDbCommand();
                cmdmat.Connection = con;
                string querymat = "SELECT stil from met";
                cmdmat.CommandText = querymat;
                OleDbDataReader readermat = cmdmat.ExecuteReader();
                while (readermat.Read())
                { comboBox4.Items.Add(readermat[0].ToString()); }

                OleDbCommand cmdceh = new OleDbCommand();
                cmdceh.Connection = con;
                string queryceh = "SELECT ceh from ceh";
                cmdceh.CommandText = queryceh;
                OleDbDataReader readerceh = cmdceh.ExecuteReader();
                while (readerceh.Read())
                { comboBox5.Items.Add(readerceh[0].ToString()); }

                OleDbCommand cmdoper = new OleDbCommand();
                cmdoper.Connection = con;
                string queryoper = "SELECT oper from oper";
                cmdoper.CommandText = queryoper;
                OleDbDataReader readeroper = cmdoper.ExecuteReader();
                while (readeroper.Read())
                { comboBox7.Items.Add(readeroper[0].ToString()); }

                OleDbCommand cmdobor = new OleDbCommand();
                cmdobor.Connection = con;
                string queryobor = "SELECT obor from obor";
                cmdobor.CommandText = queryobor;
                OleDbDataReader readerobor = cmdobor.ExecuteReader();
                while (readerobor.Read())
                { comboBox8.Items.Add(readerobor[0].ToString()); }

                OleDbCommand cmdprov = new OleDbCommand();
                cmdprov.Connection = con;
                string queryprov = "SELECT prov from prov";
                cmdprov.CommandText = queryprov;
                OleDbDataReader readerprov = cmdprov.ExecuteReader();
                while (readerprov.Read())
                { comboBox9.Items.Add(readerprov[0].ToString()); }

                OleDbCommand cmdrazr = new OleDbCommand();
                cmdrazr.Connection = con;
                string queryrazr = "SELECT razr from razr";
                cmdrazr.CommandText = queryrazr;
                OleDbDataReader readerrazr = cmdrazr.ExecuteReader();
                while (readerrazr.Read())
                { comboBox6.Items.Add(readerrazr[0].ToString()); }

                OleDbCommand cmdrinst = new OleDbCommand();
                cmdrinst.Connection = con;
                string queryrinst = "SELECT rinst from rinst";
                cmdrinst.CommandText = queryrinst;
                OleDbDataReader readerrinst = cmdrinst.ExecuteReader();
                while (readerrinst.Read())
                {
                    comboBox10.Items.Add(readerrinst[0].ToString());
                    comboBox11.Items.Add(readerrinst[0].ToString());
                    comboBox12.Items.Add(readerrinst[0].ToString());
                    comboBox13.Items.Add(readerrinst[0].ToString());
                    comboBox14.Items.Add(readerrinst[0].ToString());
                    comboBox15.Items.Add(readerrinst[0].ToString());
                }

                OleDbCommand cmdvinst = new OleDbCommand();
                cmdvinst.Connection = con;
                string queryvinst = "SELECT vinst from vinst";
                cmdvinst.CommandText = queryvinst;
                OleDbDataReader readervinst = cmdvinst.ExecuteReader();
                while (readervinst.Read())
                {
                    comboBox16.Items.Add(readervinst[0].ToString());
                    comboBox17.Items.Add(readervinst[0].ToString());
                    comboBox18.Items.Add(readervinst[0].ToString());
                    comboBox19.Items.Add(readervinst[0].ToString());
                    comboBox20.Items.Add(readervinst[0].ToString());
                    comboBox21.Items.Add(readervinst[0].ToString());
                }

                OleDbCommand cmdiinst = new OleDbCommand();
                cmdiinst.Connection = con;
                string queryiinst = "SELECT iinst from iinst";
                cmdiinst.CommandText = queryiinst;
                OleDbDataReader readeriinst = cmdiinst.ExecuteReader();
                while (readeriinst.Read())
                {
                    comboBox24.Items.Add(readeriinst[0].ToString());
                    comboBox25.Items.Add(readeriinst[0].ToString());
                    comboBox26.Items.Add(readeriinst[0].ToString());
                    comboBox27.Items.Add(readeriinst[0].ToString());
                }

                con.Close();
            }
        }


        private void админкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PUR.Form1 f2 = new PUR.Form1();
            f2.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process.Start(Path.Combine(mypath, @"SZ\"));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(Path.Combine(mypath, @"mk\"));
        }

        private void обновитьФормыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mypath = Directory.GetCurrentDirectory();
            con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Ts-03\\users\\OGK\\Кокшаров С.А\\db\\DB.accdb;Jet OLEDB:Database Password=6567604";
                        
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();
            comboBox6.Items.Clear();
            comboBox7.Items.Clear();
            comboBox8.Items.Clear();
            comboBox9.Items.Clear();
            comboBox10.Items.Clear();
            comboBox11.Items.Clear();
            comboBox12.Items.Clear();
            comboBox13.Items.Clear();
            comboBox14.Items.Clear();
            comboBox15.Items.Clear();
            comboBox16.Items.Clear();
            comboBox17.Items.Clear();
            comboBox18.Items.Clear();
            comboBox19.Items.Clear();
            comboBox20.Items.Clear();
            comboBox21.Items.Clear();
            comboBox22.Items.Clear();
            comboBox23.Items.Clear();
            comboBox24.Items.Clear();
            comboBox25.Items.Clear();
            comboBox26.Items.Clear();
            comboBox27.Items.Clear();
            
            con.Open();

            OleDbCommand cmd4 = new OleDbCommand();
            cmd4.Connection = con;
            cmd4.CommandText = "insert into stat (dat, kto) values ('" + Convert.ToString(DateTime.Now) + "','" + label25.Text + "')";
            cmd4.ExecuteNonQuery();

            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = con;
            string query = "SELECT fio1 from komu";
            cmd.CommandText = query;
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            { comboBox1.Items.Add(reader[0].ToString()); }

            OleDbCommand cmd2 = new OleDbCommand();
            cmd2.Connection = con;
            string query2 = "SELECT fio2 from ot";
            cmd2.CommandText = query2;
            OleDbDataReader reader2 = cmd2.ExecuteReader();
            while (reader2.Read())
            { comboBox2.Items.Add(reader2[0].ToString()); }

            OleDbCommand cmd3 = new OleDbCommand();
            cmd3.Connection = con;
            string query3 = "SELECT ispol from isp";
            cmd3.CommandText = query3;
            OleDbDataReader reader3 = cmd3.ExecuteReader();
            while (reader3.Read())
            { comboBox3.Items.Add(reader3[0].ToString()); }


            axAcroPDF1.LoadFile(Path.Combine(mypath, "pas.pdf"));

            //ПУ

            OleDbCommand cmdsog = new OleDbCommand();
            cmdsog.Connection = con;
            string querysog = "SELECT sog from soglas";
            cmdsog.CommandText = querysog;
            OleDbDataReader readersog = cmdsog.ExecuteReader();
            while (readersog.Read())
            { comboBox22.Items.Add(readersog[0].ToString()); }

            OleDbCommand cmdra = new OleDbCommand();
            cmdra.Connection = con;
            string queryra = "SELECT ra from sher";
            cmdra.CommandText = queryra;
            OleDbDataReader readerra = cmdra.ExecuteReader();
            while (readerra.Read())
            { comboBox23.Items.Add(readerra[0].ToString()); }

            OleDbCommand cmdmat = new OleDbCommand();
            cmdmat.Connection = con;
            string querymat = "SELECT stil from met";
            cmdmat.CommandText = querymat;
            OleDbDataReader readermat = cmdmat.ExecuteReader();
            while (readermat.Read())
            { comboBox4.Items.Add(readermat[0].ToString()); }

            OleDbCommand cmdceh = new OleDbCommand();
            cmdceh.Connection = con;
            string queryceh = "SELECT ceh from ceh";
            cmdceh.CommandText = queryceh;
            OleDbDataReader readerceh = cmdceh.ExecuteReader();
            while (readerceh.Read())
            { comboBox5.Items.Add(readerceh[0].ToString()); }

            OleDbCommand cmdoper = new OleDbCommand();
            cmdoper.Connection = con;
            string queryoper = "SELECT oper from oper";
            cmdoper.CommandText = queryoper;
            OleDbDataReader readeroper = cmdoper.ExecuteReader();
            while (readeroper.Read())
            { comboBox7.Items.Add(readeroper[0].ToString()); }

            OleDbCommand cmdobor = new OleDbCommand();
            cmdobor.Connection = con;
            string queryobor = "SELECT obor from obor";
            cmdobor.CommandText = queryobor;
            OleDbDataReader readerobor = cmdobor.ExecuteReader();
            while (readerobor.Read())
            { comboBox8.Items.Add(readerobor[0].ToString()); }

            OleDbCommand cmdprov = new OleDbCommand();
            cmdprov.Connection = con;
            string queryprov = "SELECT prov from prov";
            cmdprov.CommandText = queryprov;
            OleDbDataReader readerprov = cmdprov.ExecuteReader();
            while (readerprov.Read())
            { comboBox9.Items.Add(readerprov[0].ToString()); }

            OleDbCommand cmdrazr = new OleDbCommand();
            cmdrazr.Connection = con;
            string queryrazr = "SELECT razr from razr";
            cmdrazr.CommandText = queryrazr;
            OleDbDataReader readerrazr = cmdrazr.ExecuteReader();
            while (readerrazr.Read())
            { comboBox6.Items.Add(readerrazr[0].ToString()); }

            OleDbCommand cmdrinst = new OleDbCommand();
            cmdrinst.Connection = con;
            string queryrinst = "SELECT rinst from rinst";
            cmdrinst.CommandText = queryrinst;
            OleDbDataReader readerrinst = cmdrinst.ExecuteReader();
            while (readerrinst.Read())
            {
                comboBox10.Items.Add(readerrinst[0].ToString());
                comboBox11.Items.Add(readerrinst[0].ToString());
                comboBox12.Items.Add(readerrinst[0].ToString());
                comboBox13.Items.Add(readerrinst[0].ToString());
                comboBox14.Items.Add(readerrinst[0].ToString());
                comboBox15.Items.Add(readerrinst[0].ToString());
            }

            OleDbCommand cmdvinst = new OleDbCommand();
            cmdvinst.Connection = con;
            string queryvinst = "SELECT vinst from vinst";
            cmdvinst.CommandText = queryvinst;
            OleDbDataReader readervinst = cmdvinst.ExecuteReader();
            while (readervinst.Read())
            {
                comboBox16.Items.Add(readervinst[0].ToString());
                comboBox17.Items.Add(readervinst[0].ToString());
                comboBox18.Items.Add(readervinst[0].ToString());
                comboBox19.Items.Add(readervinst[0].ToString());
                comboBox20.Items.Add(readervinst[0].ToString());
                comboBox21.Items.Add(readervinst[0].ToString());
            }

            OleDbCommand cmdiinst = new OleDbCommand();
            cmdiinst.Connection = con;
            string queryiinst = "SELECT iinst from iinst";
            cmdiinst.CommandText = queryiinst;
            OleDbDataReader readeriinst = cmdiinst.ExecuteReader();
            while (readeriinst.Read())
            {
                comboBox24.Items.Add(readeriinst[0].ToString());
                comboBox25.Items.Add(readeriinst[0].ToString());
                comboBox26.Items.Add(readeriinst[0].ToString());
                comboBox27.Items.Add(readeriinst[0].ToString());
            }

            con.Close();
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            checkBox5.Checked = false;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            checkBox6.Checked = false;
        }
    }
}
