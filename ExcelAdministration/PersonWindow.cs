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

        public PersonWindow()
        {
            InitializeComponent();
        }

        public void loadInformation(Person person)
        {
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

        public void updateButton_Click(object sender, EventArgs e)
        {

        }

        public void saveChangesButton_Click(object sender, EventArgs e)
        {

        }

        private void newPaymentButton_Click(object sender, EventArgs e)
        {

        }

        private void cancelButton_Click(object sender, EventArgs e)
        {

        }

    }
}
