using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Geoex
{
    public partial class PrintOptions : MetroFramework.Forms.MetroForm
    {
        int closeable = 0;
        public PrintOptions(List<string> availableFields)
        {
            InitializeComponent();
            this.Opacity = 0;
            foreach (string field in availableFields)
                     chklst.Items.Add(field, true);
        }
        private void PrintOptionscs_Load(object sender, EventArgs e)
        {
            this.metroButton1.Select();
            rdoAllRows.Checked = true;
            chkFitToPageWidth.Checked = true;
            Functions.FadeIn(this, 10);
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            if (closeable < 1)
            {
                Functions.FadeOut(this, 10);
                timer1.Start();
            }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public List<string> GetSelectedColumns()
        {
            List<string> lst = new List<string>();
            foreach (object item in chklst.CheckedItems)
                lst.Add(item.ToString());
            return lst;
        }

        

        public bool PrintAllRows
        {
            get { return rdoAllRows.Checked; }
        }

        public bool FitToPageWidth
        {
            get { return chkFitToPageWidth.Checked; }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Opacity >= 0.05)
            {

            }
            else
            {
                closeable = 1;
                this.Close();
            }
        }

    }
}
