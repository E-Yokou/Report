using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateReport
{
    public partial class Report : Form
    {
        public Report()
        {
            InitializeComponent();
        }

        private void label_current_date_Click(object sender, EventArgs e)
        {
            label_current_date.Text = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToLongTimeString();
        }
    }
}
