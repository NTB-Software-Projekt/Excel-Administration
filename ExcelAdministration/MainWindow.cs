using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MemberAdministration
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void MainWindow_Load()
        {

        }

        private void exitStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void nameInput_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void otvoriExcellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel 97-2007 WorkBook|*.xls|Excel WorkBook 2010/2013|*.xlsx|All Excel Files|*.xls;*.xlsx"; //"Description|*.extension"
            ofd.ShowDialog();
            String dbPath = System.IO.Path.GetDirectoryName(ofd.FileName); //filename = System.IO.Path.GetFileName(ofd.FileName);
            textBoxPath.Text = dbPath;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AboutBox1 a = new AboutBox1();
            a.Show();
        }

    }
}
