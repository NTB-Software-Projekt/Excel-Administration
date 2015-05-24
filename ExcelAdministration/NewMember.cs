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
    public partial class NewMember : Form
    {
        private MainWindow mainForm;
        DatabaseHelper dbHelper = new DatabaseHelper();

        public NewMember(MainWindow mainForm)
        {
            InitializeComponent();
            this.mainForm = mainForm;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void addMemberButton_Click(object sender, EventArgs e)
        {
            Int32 maxID = Int32.Parse(dbHelper.maxID());
            Int32 nextID = maxID + 1;
            String ID = nextID.ToString();

            if (!String.IsNullOrEmpty(titleBox.Text) && !String.IsNullOrEmpty(surnameBox.Text) && !String.IsNullOrEmpty(nameBox.Text) && !String.IsNullOrEmpty(addressBox.Text)
                && !String.IsNullOrEmpty(zipBox.Text) && !String.IsNullOrEmpty(stateBox.Text) && !String.IsNullOrEmpty(phoneBox.Text) && !String.IsNullOrEmpty(mailBox.Text)
                && !String.IsNullOrEmpty(paymentBox.Text))
            {
                Person person = new Person(ID, titleBox.Text, surnameBox.Text, nameBox.Text, addressBox.Text, zipBox.Text, stateBox.Text, phoneBox.Text, mailBox.Text, paymentBox.Text);
                dbHelper.insertNewMember(person);
                mainForm.populateTable("");
                this.Close();
            }
            else
            {
                MessageBox.Show("Error !\nPlease fill all the required Fields !");
            }

        }
        private void resetButton_Click(object sender, EventArgs e)
        {
            titleBox.Text = "";
            surnameBox.Text = "";
            nameBox.Text = "";
            addressBox.Text = "";
            zipBox.Text = "";
            stateBox.Text = "";
            phoneBox.Text = "";
            mailBox.Text = "";
            paymentBox.Text = "";
        }
    }
}
