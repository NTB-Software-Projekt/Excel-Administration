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
        /// <summary>
        /// Initialising the GUI components by Visual Studio.
        /// </summary>
        public ExcelExample()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Closing the Window
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Window Close Event</param>
        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
