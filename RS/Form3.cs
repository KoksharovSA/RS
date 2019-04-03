using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace PUR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public string mypath;
        private void button1_Click(object sender, EventArgs e)
        {
            DateTime date1 = DateTime.Now;
            int x = date1.Day;
            int y = date1.Month;
            int z = x * y + 1002;
            if (textBox1.TextLength != 0 && textBox1.TextLength != 0 && textBox1.Text == Convert.ToString(z) && Convert.ToInt32(textBox1.Text) == z )
            {
                this.Hide();
                RS.Form2 f1 = new RS.Form2();
                f1.ShowDialog();
            }
            else
            {
                MessageBox.Show("Не попал, попробуй ещё раз.","Ты не пройдёшь!");
                return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            mypath = Directory.GetCurrentDirectory();
            pictureBox1.Image = Image.FromFile(Path.Combine(mypath, "pwd.jpg"));
            DateTime date1 = DateTime.Now;
            string d = Convert.ToString(date1.Day);
            string m = Convert.ToString(date1.Month);
            string g = Convert.ToString(date1.Year);
            label3.Text = d+"."+m+ "." + g;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
           
            if (e.KeyValue == 13) { button1_Click(sender, e); }
        }
    }
}

