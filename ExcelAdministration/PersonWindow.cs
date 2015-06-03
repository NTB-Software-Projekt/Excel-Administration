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
    /// This class manages, shows and handles user input regarding the specific information of a chosen Person.
    /// </summary>
    public partial class PersonWindow : Form
    {
        private Person person;
        DatabaseHelper dbHelper = new DatabaseHelper();
        private MainWindow mainForm;

        /// <summary>
        /// Init of all the components and specially the mainForm which is needed for the refreshing of the member table upon closing of this Window.
        /// </summary>
        /// <param name="mainForm">The mainForm window passed by reference for refreshing the table</param>
        public PersonWindow(MainWindow mainForm)
        {
            this.mainForm = mainForm;
            InitializeComponent();
        }

        /// <summary>
        /// Loading the Text in the TextBoxes of the GUI.
        /// </summary>
        /// <param name="person">Datamodel passed by reference containing all needed data</param>
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

        /// <summary>
        /// Enables the TexBoxes so that the Data can be edited and later updated
        /// </summary>
        /// /// <param name="sender">Button</param>
        /// <param name="e">Function Call Event</param>
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

        /// <summary>
        /// Packages all the new Data into the Datamodel Person and saves it in the Database.
        /// Upon saving the Data it disables the textboxes again and refreshes the table in the Mainfrom;
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Function Call Event</param>
        public void saveChangesButton_Click(object sender, EventArgs e)
        {
            Person toUpdate = new Person(person.ID, titleBox.Text, surnameBox.Text, nameBox.Text, addressBox.Text, zipBox.Text, stateBox.Text, phoneBox.Text, emailBox.Text, payedBox.Text);
            dbHelper.updateMember(toUpdate);
            disableTextBoxes();
            mainForm.populateTable("");
        }

        /// <summary>
        /// Adds the inputed payment to the already payed ammount from the Database and updates the TextBox with the new Data.
        /// Doesn't save anything to the Database. 
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Function Call Event</param>
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

        /// <summary>
        /// Disables the TextBoxes so nothing can be written.
        /// </summary>
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

        /// <summary>
        /// Deletes the chosen Person from the Database*.
        /// As the driver doesn't support the direct deletition of Entries in the Excel File, the solution to this was sending a "null" update to the Excel File
        /// where all columns in this row are set to null.
        /// This method has a bug. If we were to delete the last entry in the Table all function of the application stop working.
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Function Call Event</param>
        private void deleteButton_Click(object sender, EventArgs e)
        {
            dbHelper.deleteMember(person.ID);
            mainForm.populateTable("");
            this.Close();
        }

    }
}
