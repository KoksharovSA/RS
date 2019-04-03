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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public string mypath;
        private void Form2_Load(object sender, EventArgs e)
        {

            this.statTableAdapter.Fill(this.dBDataSet.stat);
            this.loginTableAdapter.Fill(this.dBDataSet.login);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dBDataSet.login". При необходимости она может быть перемещена или удалена.
            this.loginTableAdapter.Fill(this.dBDataSet.login);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dBDataSet.login". При необходимости она может быть перемещена или удалена.
            this.loginTableAdapter.Fill(this.dBDataSet.login);

            this.pripuskiTableAdapter.Fill(this.dBDataSet1.pripuski);
            mypath = Directory.GetCurrentDirectory();

            this.provTableAdapter.Fill(this.dBDataSet.prov);
            this.razrTableAdapter.Fill(this.dBDataSet.razr);
            this.iinstTableAdapter.Fill(this.dBDataSet.iinst);
            this.vinstTableAdapter.Fill(this.dBDataSet.vinst);
            this.rinstTableAdapter.Fill(this.dBDataSet.rinst);
            this.cehTableAdapter.Fill(this.dBDataSet.ceh);
            this.oborTableAdapter.Fill(this.dBDataSet.obor);
            this.metTableAdapter.Fill(this.dBDataSet.met);
            this.operTableAdapter.Fill(this.dBDataSet.oper);
            this.ispTableAdapter.Fill(this.dBDataSet.isp);
            this.otTableAdapter.Fill(this.dBDataSet.ot);
            this.komuTableAdapter.Fill(this.dBDataSet.Komu);
            this.bDRTableAdapter.Fill(this.dBDataSet.BDR);
            
               
            
        }        

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.provTableAdapter.Update(this.dBDataSet.prov);
            this.razrTableAdapter.Update(this.dBDataSet.razr);
            this.iinstTableAdapter.Update(this.dBDataSet.iinst);
            this.vinstTableAdapter.Update(this.dBDataSet.vinst);
            this.rinstTableAdapter.Update(this.dBDataSet.rinst);
            this.cehTableAdapter.Update(this.dBDataSet.ceh);
            this.oborTableAdapter.Update(this.dBDataSet.obor);
            this.metTableAdapter.Update(this.dBDataSet.met);
            this.operTableAdapter.Update(this.dBDataSet.oper);
            this.ispTableAdapter.Update(this.dBDataSet.isp);
            this.otTableAdapter.Update(this.dBDataSet.ot);
            this.komuTableAdapter.Update(this.dBDataSet.Komu);
            this.bDRTableAdapter.Update(this.dBDataSet.BDR);
            this.statTableAdapter.Update(this.dBDataSet.stat);
            this.loginTableAdapter.Update(this.dBDataSet.login);
            this.pripuskiTableAdapter.Update(this.dBDataSet1.pripuski);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Файл локальной базы данных будет заменён на файл сетевой базы данных и восстановить его будет невозможно. Ты уверен?", "Обновление локальной базы данных", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string put1 = @"\\Ts-03\users\OGK\Кокшаров С.А\db\DB.accdb";
                string put2 = Path.Combine(mypath, "DB.accdb");
                File.Copy(put1, put2, true);
                MessageBox.Show("Файл ЛБД заменён на файл ЛБД", "Успешно");
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            string put1 = @"\\Ts-03\users\OGK\Кокшаров С.А\db\DB.accdb";
            string put2 = @"\\Ts-03\users\OGK\Кокшаров С.А\db\backup\DB_" + DateTime.Now.Date.ToString("dd.MM.yyyy") + ".accdb";
            File.Copy(put1, put2, true);
            MessageBox.Show("Backup СБД создан(DB_" + DateTime.Now.Date.ToString("dd.MM.yyyy") + ".accdb" + ") ", "Backup СБД");
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
