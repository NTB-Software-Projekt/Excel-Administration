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
    /// Simple example how a Excel File should be formatted for this application to work.
    /// It only serves as Help so there is no additional functionality.
    /// </summary>
    public partial class ExcelExample : Form
    {
        public ExcelExample()
        {
            InitializeComponent();
        }

        private void ExcelExample_Load(object sender, EventArgs e)
        {

        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
