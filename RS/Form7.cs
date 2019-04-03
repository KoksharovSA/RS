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

namespace RS
{
    public partial class Form7 : Form
    {
        public Form7()
        {
            InitializeComponent();
        }

        private OleDbConnection con = new OleDbConnection();
        public string mypath;
        private void Form7_Load(object sender, EventArgs e)
        {
            mypath = Directory.GetCurrentDirectory();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            DateTime date1 = DateTime.Now;
            int x = date1.Day;
            int y = date1.Month;
            int z = x * y + 1001;
            try
            {
                if (textBox1.Text == "admin" & textBox2.Text == Convert.ToString(z))
                {
                    this.Hide();
                    RS.Form1 f1 = new RS.Form1();
                    f1.label25.Text = "admin";
                    f1.ShowDialog();
                }
                con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Ts-03\\users\\OGK\\Кокшаров С.А\\db\\DB.accdb;Jet OLEDB:Database Password=6567604";
                con.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                string query = "SELECT namelogin, password, m, y from login";
                cmd.CommandText = query;
                OleDbDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    if (textBox1.Text == reader[0].ToString() & textBox2.Text == reader[1].ToString())
                    {
                        int d1y = date1.Year;
                        int d1m = date1.Month;
                        int d2 = Convert.ToInt32(reader[3].ToString()) + 1987;
                        int d3 = Convert.ToInt32(reader[2].ToString()) - 19;
                        if (d1y < d2)
                        {
                            this.Hide();
                            RS.Form1 f1 = new RS.Form1();
                            f1.label25.Text = reader[0].ToString();
                            f1.ShowDialog();
                            break;
                        }
                        if (d1y == d2)
                        {
                            if (d1m <= d3)
                            {
                                if (d1m < d3)
                                {
                                    this.Hide();
                                    RS.Form1 f1 = new RS.Form1();
                                    f1.label25.Text = reader[0].ToString();
                                    f1.ShowDialog();
                                    break;
                                }
                                if (d1m == d3)
                                {
                                    MessageBox.Show("Программа будет запущена, но срок лицензии истечёт в следующем месяце))");
                                    this.Hide();
                                    RS.Form1 f1 = new RS.Form1();
                                    f1.label25.Text = reader[0].ToString();
                                    f1.ShowDialog();
                                    break;
                                }
                            }
                            else { MessageBox.Show("Срок лицензии истёк"); }
                        }
                        else { MessageBox.Show("Срок лицензии истёк"); }
                    }                    
                }
                con.Close();
            }
            catch
            {
                con.Close();
                MessageBox.Show("Что-то пошло не так", "Ошибка");               
            }
        }

        private void textBox2_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyValue == 13) { button1_Click(sender, e); }
        }
    }
}
