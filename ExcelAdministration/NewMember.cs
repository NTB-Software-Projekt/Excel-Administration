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
        MainWindow form;

        public NewMember(MainWindow form)
        {
            InitializeComponent();
            this.form = form;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void addMemberButton_Click(object sender, EventArgs e)
        {

        }
        private void resetButton_Click(object sender, EventArgs e)
        {

        }
    }
}
