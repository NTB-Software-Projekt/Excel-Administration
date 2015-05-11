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
                MessageBox.Show("No Database Selected");
            }
            else 
            {
                dbHelper.setDatabase(MemberAdministration.Properties.Settings.Default.dbPath);
                writeTextBoxPath(MemberAdministration.Properties.Settings.Default.dbPath);
                populateTable();
            }
        }

        private void exitStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void nameInput_TextChanged(object sender, EventArgs e)
        {
            String inputName = textBox1.Text;
            populateTable(inputName);
        }

        private void openExcelMenu_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "All Excel Files|*.xls;*.xlsx|Excel 97-2007 WorkBook|*.xls|Excel WorkBook 2010/2013|*.xlsx"; //"Description|*.extension"
            ofd.ShowDialog();
            String dbPath = System.IO.Path.GetDirectoryName(ofd.FileName)+"\\"+System.IO.Path.GetFileName(ofd.FileName);
            writeTextBoxPath(dbPath);

            dbHelper.setDatabase(dbPath);
            populateTable();

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

        public void populateTable()
        {
            BindingSource source = new BindingSource();
            source.DataSource = dbHelper.getAllMembers().Tables[0];
            dataView.DataSource = source;
            DataGridViewColumn column = dataView.Columns[0];
            column.Width = 40;
        }

        public void populateTable(String name)
        {
            BindingSource source = new BindingSource();
            try
            {
                source.DataSource = dbHelper.getThisMember(name).Tables[0];
                dataView.DataSource = source;

                DataGridViewColumn column = dataView.Columns[0];
                column.Width = 40;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

    }
}
