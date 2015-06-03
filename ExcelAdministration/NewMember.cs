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
    /// Handles the Addition of new members trough a form.
    /// </summary>
    public partial class NewMember : Form
    {
        private MainWindow mainForm;
        DatabaseHelper dbHelper = new DatabaseHelper();

        /// <summary>
        /// Init of all the components and specially the mainForm which is needed for the refreshing of the member table upon closing of this Window.
        /// </summary>
        /// <param name="mainForm">The mainForm window passed by reference for refreshing the table</param>
        public NewMember(MainWindow mainForm)
        {
            InitializeComponent();
            this.mainForm = mainForm;
        }

        /// <summary>
        /// Closes the Window upon cancellation.
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Window Close Event</param>
        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Adding a new Member to the Database with all the Information in the Form.
        /// If the Database is empty it will add a new Member in the first position, 
        /// else it will search the last added ID and add a new member with the last ID + 1.
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Function Call Event</param>
        private void addMemberButton_Click(object sender, EventArgs e)
        {
            String ID;

            if (!String.IsNullOrEmpty(dbHelper.maxID()))
            {
                Int32 maxID = Int32.Parse(dbHelper.maxID());
                Int32 nextID = maxID + 1;
                ID = nextID.ToString();
            }
            else
            {
                ID = "1";
            }
            
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

        /// <summary>
        /// Clears the TextBoxes.
        /// </summary>
        /// /// <param name="sender">Button</param>
        /// <param name="e">Window Event</param>
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
