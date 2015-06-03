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
    /// <summary>
    /// This is the Core Class. It takes the input from the user and handles all calls to open subwindows,
    /// and generate backups.
    /// </summary>
    public partial class MainWindow : Form
    {
        DatabaseHelper dbHelper = new DatabaseHelper();
        BackupManager backupMannager = new BackupManager();
        PersonWindow pWindow;

        /// <summary>
        /// Upon creating the Main Window the backupManager saves the last state of the Excel File.
        /// When something goes wrong the old Excel File is then secured in the backup folder.
        /// </summary>
        public MainWindow()
        {
            backupMannager.saveFile();
            InitializeComponent();
        }

        /// <summary>
        /// If a Database (Excel File) has been chosen in a previous session it will be loaded from the application properties.
        /// Otherwise it will alert the user to choose one.
        /// </summary>
        /// <param name="sender">Form Load</param>
        /// <param name="e">Load Event</param>
        public void MainWindow_Load(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(MemberAdministration.Properties.Settings.Default.dbPath))
            {
                MessageBox.Show("First Time starting the Application.\nWelcome to xl2DB \nPlease choose a Database from the Menu");
            }
            else
            {
                dbHelper.setDatabase(MemberAdministration.Properties.Settings.Default.dbPath);
                writeTextBoxPath(MemberAdministration.Properties.Settings.Default.dbPath);
                populateTable("");
            }
        }

        /// <summary>
        /// Exiting the Application trough the Menu strip.
        /// </summary>
        /// <param name="sender">Menu strip Click</param>
        /// <param name="e">Close Event</param>
        private void exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// On every text change a SQL query is sent to populate the table with the Text from the Textbox.
        /// </summary>
        /// <param name="sender">Text Label</param>
        /// <param name="e">Text Changed Event</param>
        private void nameInput_TextChanged(object sender, EventArgs e)
        {
            String inputName = inputBox.Text;
            populateTable(inputName);
        }

        /// <summary>
        /// Opens the Filechooser window so that the Excel File needed as Database can be selected.
        /// </summary>
        /// <param name="sender">Menu Strip</param>
        /// <param name="e">Window Open Event</param>
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

        /// <summary>
        /// Opens the About us box with the project information.
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Window Open Event</param>
        private void aboutUs_Click(object sender, EventArgs e)
        {
            AboutBox a = new AboutBox();
            a.Show();
        }

        /// <summary>
        /// Writes the Path to the current Database on the GUI in the right top corner.
        /// </summary>
        /// <param name="dbPath">Path of the current/selected Database</param>
        private void writeTextBoxPath(String dbPath)
        {
            textBoxPath.Text = dbPath;
        }

        /// <summary>
        /// Populates the Table from the database with all members
        /// </summary>
        public void populateTable()
        {
            BindingSource source = new BindingSource();
            source.DataSource = dbHelper.getAllMembers().Tables[0];
            dataView.DataSource = source;
            dataView.Columns[0].Visible = false;

            DataGridViewColumn column = dataView.Columns[1];
            column.Width = 50;
        }

        /// <summary>
        /// Populates the table from the database with the member that matches the name, surname, address or state with the parameter.
        /// </summary>
        /// <param name="name">String which we want to match in the database</param>
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

        /// <summary>
        /// Opens the new member form
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Window Open Event</param>
        private void newMemberButton_Click(object sender, EventArgs e)
        {
            NewMember newMember = new NewMember(this);
            newMember.Show();
        }

        /// <summary>
        /// Opens the Person Window with all the needed information from the Table.
        /// </summary>
        /// <param name="sender">Cell Click</param>
        /// <param name="e">Clicked Cell</param>
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

        /// <summary>
        /// Refreshes the table, hiding the empty entries from the Excel File
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">GridView Load Event</param>
        private void refreshButton_Click(object sender, EventArgs e)
        {
            populateTable("");
        }

        /// <summary>
        /// Opens the Excel Example form which only shows how the Excel File needs to be formated.
        /// </summary>
        /// <param name="sender">Menu Strip</param>
        /// <param name="e">Window Open Event</param>
        private void exampleMenu_Click(object sender, EventArgs e)
        {
            ExcelExample exampleBox = new ExcelExample();
            exampleBox.Show();
        }

    }
}
