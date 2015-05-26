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
        BackupManager backupMannager = new BackupManager();
        PersonWindow pWindow;

        public MainWindow()
        {
            backupMannager.saveFile();
            InitializeComponent();
        }

        public void MainWindow_Load(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(MemberAdministration.Properties.Settings.Default.dbPath))
            {
                MessageBox.Show("First Time starting the Application.\nWelcome to xl2DB");
            }
            else
            {
                dbHelper.setDatabase(MemberAdministration.Properties.Settings.Default.dbPath);
                writeTextBoxPath(MemberAdministration.Properties.Settings.Default.dbPath);
                populateTable("");
            }
        }

        private void exit_Click(object sender, EventArgs e)
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
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                String dbPath = System.IO.Path.GetDirectoryName(ofd.FileName) + "\\" + System.IO.Path.GetFileName(ofd.FileName);
                writeTextBoxPath(dbPath);
                backupMannager.setFileName(System.IO.Path.GetFileName(ofd.FileName));
                dbHelper.setDatabase(dbPath);
                populateTable("");
            }
        }

        private void aboutUs_Click(object sender, EventArgs e)
        {
            AboutBox1 a = new AboutBox1();
            a.Show();
        }

        private void writeTextBoxPath(String dbPath)
        {
            textBoxPath.Text = dbPath;
        }

        public void populateTable()
        {
            BindingSource source = new BindingSource();
            source.DataSource = dbHelper.getAllMembers().Tables[0];
            dataView.DataSource = source;
            dataView.Columns[0].Visible = false;

            DataGridViewColumn column = dataView.Columns[1];
            column.Width = 50;
        }

        public void populateTable(String name)
        {
            BindingSource source = new BindingSource();
            try
            {
                source.DataSource = dbHelper.getThisMember(name).Tables[0];
                dataView.DataSource = source;
                dataView.Columns[0].Visible = false;

                DataGridViewColumn column = dataView.Columns[1];
                column.Width = 50;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void newMemberButton_Click(object sender, EventArgs e)
        {
            NewMember newMember = new NewMember(this);
            newMember.Show();
        }

        private void cellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (pWindow != null)
            {
                pWindow.Close();
            }

            if (e.RowIndex >= 0)
            {
                String ID = "";
                String title = "";
                String surname = "";
                String name = "";
                String address = "";
                String zip = "";
                String state = "";
                String telephone = "";
                String email = "";
                String amount = "";

                foreach (DataGridViewRow row in dataView.SelectedRows)
                {
                    ID = row.Cells[0].Value.ToString();
                    title = row.Cells[1].Value.ToString();
                    surname = row.Cells[2].Value.ToString();
                    name = row.Cells[3].Value.ToString();
                    address = row.Cells[4].Value.ToString();
                    zip = row.Cells[5].Value.ToString();
                    state = row.Cells[6].Value.ToString();
                    telephone = row.Cells[7].Value.ToString();
                    email = row.Cells[8].Value.ToString();
                    amount = row.Cells[9].Value.ToString();
                }

                Person person = new Person(ID, title, surname, name, address, zip, state, telephone, email, amount);
                pWindow = new PersonWindow(this);
                pWindow.loadInformation(person);
                pWindow.Show();
            }
        }

        private void refreshButton_Click(object sender, EventArgs e)
        {
            populateTable("");
        }

        private void exampleMenu_Click(object sender, EventArgs e)
        {
            ExcelExample exampleBox = new ExcelExample();
            exampleBox.Show();
        }



    }
}
