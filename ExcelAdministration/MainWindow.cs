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
        DatabaseHelper dbHelper = new DatabaseHelper();
        public MainWindow()
        {
            InitializeComponent();
        }
        
        public void MainWindow_Load(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(MemberAdministration.Properties.Settings.Default.dbPath))
            {
                MessageBox.Show("No Database");
            }
            else 
            {
                dbHelper.setDatabase(MemberAdministration.Properties.Settings.Default.dbPath);
                writeTextBoxPath(MemberAdministration.Properties.Settings.Default.dbPath);
            }
        }

        private void exitStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void nameInput_TextChanged(object sender, EventArgs e)
        {
            //TODO
        }

        private void openExcelMenu_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel 97-2007 WorkBook|*.xls|Excel WorkBook 2010/2013|*.xlsx|All Excel Files|*.xls;*.xlsx"; //"Description|*.extension"
            ofd.ShowDialog();
            String dbPath = System.IO.Path.GetDirectoryName(ofd.FileName); //filename = System.IO.Path.GetFileName(ofd.FileName);
            writeTextBoxPath(dbPath);

            dbHelper.setDatabase(dbPath);

        }

        private void writeTextBoxPath(String dbPath)
        {
            textBoxPath.Text = dbPath;
        }

        private void exitApplicationMenu_Click(object sender, EventArgs e)
        {
            AboutBox1 a = new AboutBox1();
            a.Show();
        }

    }
}
