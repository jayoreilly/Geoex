using Geoex.Properties;
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
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        public Form1()
        {
            InitializeComponent();
            this.Opacity = 0;
        }


        int closeable = 0;
        private void Form1_Load(object sender, EventArgs e)
        {
            Functions.FadeIn(this, 10);
  
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel1.Text))//GARDNER DENVER 2000
                {
                    box.ShowDialog();
                }  
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (closeable < 1)
            {
                e.Cancel = true;
                Functions.FadeOut(this, 10);
                timer1.Start();
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel2.Text))//GARDNER DENVER 15W
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel2.Text))//GARDNER DENVER 15W
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel5.Text))//MAYHEW 1000
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel3.Text))//MIDWAY 15M
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel4.Text))//IR CYCLON RO 300
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel14.Text))//BUCYRUS ERIE 22
            {
                box.ShowDialog();
            }  
            
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel10.Text))//BUCYRUS ERIE 36L
            {
                box.ShowDialog();
            }  
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Opacity >= 0.05)
            {
                
            }
            else
            {
                closeable = 1;
                Application.Exit();
            }
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel15.Text))//INGERSOLL RAND (1)
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel13.Text))//INGERSOLL RAND (2)
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel12.Text))//SULLAIR
            {
                box.ShowDialog();
            }  
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1(metroLabel27.Text))//NCA
            {
                box.ShowDialog();
            }  
        }
    }
}
