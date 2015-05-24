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
    public partial class PersonWindow : Form
    {
        private Person person;
        DatabaseHelper dbHelper = new DatabaseHelper();
        private MainWindow mainForm;

        public PersonWindow(MainWindow mainForm)
        {
            this.mainForm = mainForm;
            InitializeComponent();
        }

        public void loadInformation(Person person)
        {
            this.person = person;
            titleBox.Text = person.Title;
            surnameBox.Text = person.Surname;
            nameBox.Text = person.Name;
            addressBox.Text = person.Address;
            zipBox.Text = person.Plz;
            stateBox.Text = person.State;
            phoneBox.Text = person.Telephone.ToString();
            emailBox.Text = person.Mail;
            payedBox.Text = person.Amount.ToString();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            titleBox.Enabled = true;
            surnameBox.Enabled = true;
            nameBox.Enabled = true;
            addressBox.Enabled = true;
            zipBox.Enabled = true;
            stateBox.Enabled = true;
            phoneBox.Enabled = true;
            emailBox.Enabled = true;
            payedBox.Enabled = true;
        }

        public void saveChangesButton_Click(object sender, EventArgs e)
        {
            Person toUpdate = new Person(person.ID, titleBox.Text, surnameBox.Text, nameBox.Text, addressBox.Text, zipBox.Text, stateBox.Text, phoneBox.Text, emailBox.Text, payedBox.Text);
            dbHelper.updateMember(toUpdate);
            disableTextBoxes();
            mainForm.populateTable("");
        }

        private void newPaymentButton_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(newPaymentBox.Text))
            {
                Int32 alreadyPayed = Int32.Parse(payedBox.Text);
                Int32 newPayment = Int32.Parse(newPaymentBox.Text);

                payedBox.Text = (alreadyPayed + newPayment).ToString();
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            disableTextBoxes();
        }

        private void disableTextBoxes()
        {
            titleBox.Enabled = false;
            surnameBox.Enabled = false;
            nameBox.Enabled = false;
            addressBox.Enabled = false;
            zipBox.Enabled = false;
            stateBox.Enabled = false;
            phoneBox.Enabled = false;
            emailBox.Enabled = false;
            payedBox.Enabled = false;
        }

        private void deleteButton_Click(object sender, EventArgs e)
        {
            dbHelper.deleteMember(person.ID);
            mainForm.populateTable("");
            this.Close();
        }

    }
}
