namespace Geoex
{
    using Geoex.Properties;
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Drawing.Printing;
    using System.IO;
    using System.Reflection;
    using System.Windows.Forms;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Security.Permissions;
    using System.Data.OleDb;

    public partial class AboutBox1 : MetroFramework.Forms.MetroForm
    {
        public string dir;
        public int rowIndex = 0;
        public string lb;
        int closeable = 0;

        public AboutBox1(string lbt)
        {
            this.lb = lbt;
            this.InitializeComponent();
            this.Opacity = 0;
            DateTime dtCurrentTime = DateTime.Now;
           Time.Text = dtCurrentTime.ToLongTimeString();
           timer2.Start();
        }
        
        private void AboutBox1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.listView1.Visible == false)
            {
                Functions.FadeOut(this, 10);            //Do you wish to save your changes
                DialogResult result = MessageBox.Show("Desea guardar los cambios?", "Geoex", MessageBoxButtons.YesNoCancel);

                if (result == DialogResult.Yes)
                {
                    savenotes();
                    if (el1.Text == "GARDNER DENVER 2000")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "GARDNER DENVER 15W")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "MIDWAY 15M")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "MAYHEW 1000")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "IR CYCLON RO 300")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "INGERSOLL RAND (2)")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "INGERSOLL RAND (1)")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "BUCYRUS ERIE 36L")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "BUCYRUS ERIE 22")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "SULLAIR")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                    else if (el1.Text == "NCA")
                    {
                        Functions.SaveasClose(this.dgv, dir);
                    }
                
                }
                if (result == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    Functions.FadeIn(this, 10);
                }
            }
            else
            {
                if (closeable < 1)
                {
                    Functions.FadeOut(this, 10);
                    e.Cancel = true;
                    Timer.Start();
                }
            }
        }

        public void Calculate()
        {
            try
            {
                
                object a = dgv.CurrentRow.Cells[3].Value; object b = dgv.CurrentRow.Cells[4].Value;
                object c = dgv.CurrentRow.Cells[5].Value; object d = dgv.CurrentRow.Cells[6].Value;
                object ee = dgv.CurrentRow.Cells[7].Value; object f = dgv.CurrentRow.Cells[8].Value;
                object g = dgv.CurrentRow.Cells[9].Value; object h = dgv.CurrentRow.Cells[10].Value;
                object i = dgv.CurrentRow.Cells[11].Value; object j = dgv.CurrentRow.Cells[12].Value;
                object k = dgv.CurrentRow.Cells[13].Value; object l = dgv.CurrentRow.Cells[14].Value;
                double aNumber = 0; double bNumber = 0;
                double cNumber = 0; double dNumber = 0;
                double eeNumber = 0; double fNumber = 0;
                double gNumber = 0; double hNumber = 0;
                double iNumber = 0; double jNumber = 0;
                double kNumber = 0; double lNumber = 0;

                if (a != DBNull.Value)
                    aNumber = Double.Parse(a.ToString());
                if (b != DBNull.Value)
                    bNumber = Double.Parse(b.ToString());
                if (c != DBNull.Value)
                    cNumber = Double.Parse(c.ToString());
                if (d != DBNull.Value)
                    dNumber = Double.Parse(d.ToString());
                if (ee != DBNull.Value)
                    eeNumber = Double.Parse(ee.ToString());
                if (f != DBNull.Value)
                    fNumber = Double.Parse(f.ToString());
                if (g != DBNull.Value)
                    gNumber = Double.Parse(g.ToString());
                if (h != DBNull.Value)
                    hNumber = Double.Parse(h.ToString());
                if (i != DBNull.Value)
                    iNumber = Double.Parse(i.ToString());
                if (j != DBNull.Value)
                    jNumber = Double.Parse(j.ToString());
                if (k != DBNull.Value)
                    kNumber = Double.Parse(k.ToString());
                if (l != DBNull.Value)
                    lNumber = Double.Parse(l.ToString());
                dgv.CurrentRow.Cells[2].Value = aNumber + bNumber + cNumber + dNumber + eeNumber + fNumber + gNumber + hNumber + iNumber + jNumber + kNumber + lNumber;
            }
            catch {}
            
            }

        private void AboutBox1_Load(object sender, EventArgs e)
        {
            Notes.Visible = false;
            saveasToolStripButton.Enabled = false;
            saveToolStripButton.Enabled = false;
            printToolStripButton.Enabled = false;
            Functions.FadeIn(this, 10);
            el1.Text = lb;
            this.listView1.Visible = true;
            this.metroLabel1.Visible = true;
            this.dgv.Visible = false;
            this.Size = new Size(380, 620);
             ListViewItem item;

             try
             {

                 if (el1.Text == "GARDNER DENVER 2000")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\GARDNER DENVER 2000\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }

                  
                 }
                 else if (el1.Text == "GARDNER DENVER 15W")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\GARDNER DENVER 15W\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "MIDWAY 15M")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\MIDWAY 15M\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "NCA")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\NCA\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "SULLAIR")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\SULLAIR\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "BUCYRUS ERIE 22")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\BUCYRUS ERIE 22\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "BUCYRUS ERIE 36L")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\BUCYRUS ERIE 36L\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "INGERSOLL RAND (1)")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\INGERSOLL RAND (1)\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "INGERSOLL RAND (2)")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\INGERSOLL RAND (2)\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "IR CYCLON RO 300")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\IR CYCLON RO 300\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
                 else if (el1.Text == "MAYHEW 1000")
                 {

                     string directoryName = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                     this.listView1.SmallImageList = Icons;

                     string currentDirectory = directoryName;
                     string[] directories = Directory.GetDirectories(currentDirectory);
                     int length = directories.Length;
                     string[] files = Directory.GetFiles(currentDirectory + @"\MAYHEW 1000\", "*.xlsx");
                     length = files.Length;
                     for (int j = 0; j < length; j++)
                     {
                         item = new ListViewItem(Path.GetFileName(files[j]), 0);
                         FileInfo info = new FileInfo(files[j]);
                         this.listView1.Items.Add(item);
                     }
                 }
             }



             catch (Exception exception)
             {
                 MessageBox.Show(exception.Message);
             }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            Info.Icon = SystemIcons.Information;
            Info.BalloonTipText = "Cargando. Favor de esperar"; 
            Info.ShowBalloonTip(2000);
            String text = listView1.SelectedItems[0].Text;
            loadnotes();
             try
            {
                if (el1.Text == "GARDNER DENVER 2000")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\GARDNER DENVER 2000\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "GARDNER DENVER 15W")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\GARDNER DENVER 15W\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "MIDWAY 15M")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\MIDWAY 15M\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "MAYHEW 1000")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\MAYHEW 1000\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "IR CYCLON RO 300")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\IR CYCLON RO 300\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "INGERSOLL RAND (2)")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\INGERSOLL RAND (2)\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "INGERSOLL RAND (1)")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\INGERSOLL RAND (1)\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "BUCYRUS ERIE 36L")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\BUCYRUS ERIE 36L\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "BUCYRUS ERIE 22")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\BUCYRUS ERIE 22\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "SULLAIR")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\SULLAIR\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
                else if (el1.Text == "NCA")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\NCA\" + text;
                    Functions.OpenedOnFormLoad(path, dgv, bS);
                    dir = path;
                }
            }
            catch
            {
                
            }
            this.dgv.Visible = true;
            this.listView1.Visible = false;
            this.metroLabel1.Visible = false;
            this.WindowState = FormWindowState.Maximized;
            this.toolStripStatusLabel1.Text = "Exitosamente cargado " + text;
            saveasToolStripButton.Enabled = true;
            saveToolStripButton.Enabled = true;
            printToolStripButton.Enabled = true;
            Notes.Visible = true;
        }

        private void newToolStripButton_Click(object sender, EventArgs e)
        {
            loadnotes();
            try
            {
                if (el1.Text == "GARDNER DENVER 2000")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\GARDNER DENVER 2000\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "GARDNER DENVER 15W")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\GARDNER DENVER 15W\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "MIDWAY 15M")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\MIDWAY 15M\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "MAYHEW 1000")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\MAYHEW 1000\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "IR CYCLON RO 300")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\IR CYCLON RO 300\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "INGERSOLL RAND (2)")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\INGERSOLL RAND (2)\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "INGERSOLL RAND (1)")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\INGERSOLL RAND (1)\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "BUCYRUS ERIE 36L")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\BUCYRUS ERIE 36L\";
                    Functions.NewFile(this.dgv, this.bS, path); 
                }
                else if (el1.Text == "BUCYRUS ERIE 22")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\BUCYRUS ERIE 22\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "SULLAIR")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\SULLAIR\";
                    Functions.NewFile(this.dgv, this.bS, path);
                }
                else if (el1.Text == "NCA")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Templates\NCA\";
                    Functions.NewFile(this.dgv, this.bS, path);
                    Notes.Text = "NCA";
                }
                
            }
            catch
            {
            }
            if (Functions.OpenV == true)
            {
                this.dgv.Visible = true;
                this.listView1.Visible = false;
                this.metroLabel1.Visible = false;
                this.WindowState = FormWindowState.Maximized;
                this.toolStripStatusLabel1.Text = "Exitosamente cargado ";
                saveasToolStripButton.Enabled = true;
                saveToolStripButton.Enabled = true;
                printToolStripButton.Enabled = true;
                Notes.Visible = true;
                Info.Icon = SystemIcons.Information;
                Info.BalloonTipText = "Cargando. Favor de esperar";
                Info.ShowBalloonTip(2000);
            }
            else
            {

            }
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {       loadnotes();
            try
            {  
                    if (el1.Text == "GARDNER DENVER 2000")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\GARDNER DENVER 2000\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "GARDNER DENVER 15W")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\GARDNER DENVER 15W\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "MIDWAY 15M")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\MIDWAY 15M\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "MAYHEW 1000")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\MAYHEW 1000\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "IR CYCLON RO 300")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\IR CYCLON RO 300\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "INGERSOLL RAND (2)")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\INGERSOLL RAND (2)\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "INGERSOLL RAND (1)")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\INGERSOLL RAND (1)\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "BUCYRUS ERIE 36L")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\BUCYRUS ERIE 36L\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "BUCYRUS ERIE 22")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\BUCYRUS ERIE 22\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "SULLAIR")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\SULLAIR\";
                        Functions.Open(this.dgv, this.bS, path);
                    }
                    else if (el1.Text == "NCA")
                    {
                        string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\NCA\";
                        Functions.Open(this.dgv, this.bS, path);
                    }

                
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            if (Functions.OpenV == true)
            {
                this.dgv.Visible = true;
                this.listView1.Visible = false;
                this.metroLabel1.Visible = false;
                this.WindowState = FormWindowState.Maximized;
                saveasToolStripButton.Enabled = true;
                saveToolStripButton.Enabled = true;
                printToolStripButton.Enabled = true;
                Notes.Visible = true;
                Info.Icon = SystemIcons.Information;
                Info.BalloonTipText = "Cargando. Favor de esperar";
                Info.ShowBalloonTip(2000);
                this.toolStripStatusLabel1.Text = "Exitosamente cargado " + Functions.SafeFileName();
            }
            else
            {
              
            }
        }

        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                PrintDGV.Print_DataGridView(this.dgv);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void saveasToolStripButton_Click(object sender, EventArgs e)
        {
            savenotes();
            try
            {
                
                if (el1.Text == "GARDNER DENVER 2000")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\GARDNER DENVER 2000\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "GARDNER DENVER 15W")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\GARDNER DENVER 15W\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "MIDWAY 15M")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\MIDWAY 15M\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "MAYHEW 1000")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\MAYHEW 1000\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "IR CYCLON RO 300")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\IR CYCLON RO 300\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "INGERSOLL RAND (2)")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\INGERSOLL RAND (2)\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "INGERSOLL RAND (1)")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\INGERSOLL RAND (1)\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "BUCYRUS ERIE 36L")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\BUCYRUS ERIE 36L\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "BUCYRUS ERIE 22")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\BUCYRUS ERIE 22\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "SULLAIR")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\SULLAIR\";
                    Functions.Saveas(this.dgv, path);
                }
                else if (el1.Text == "NCA")
                {
                    string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\NCA\";
                    Functions.Saveas(this.dgv, path);
                }
                if (Functions.OpenV == true)
                {
                    Info.Icon = SystemIcons.Information;
                    Info.BalloonTipText = "Guardando. Favor de esperar";
                    Info.ShowBalloonTip(2000);
                    this.toolStripStatusLabel1.Text = Functions.SafeFileName() + " Cargado exitosamente el " + DateTime.Now.ToShortTimeString();
                }
                else
                {

                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            savenotes();
            
                try
                {
                    if (el1.Text == "GARDNER DENVER 2000")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "GARDNER DENVER 15W")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "MIDWAY 15M")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "MAYHEW 1000")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "IR CYCLON RO 300")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "INGERSOLL RAND (2)")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "INGERSOLL RAND (1)")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "BUCYRUS ERIE 36L")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "BUCYRUS ERIE 22")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "SULLAIR")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    else if (el1.Text == "NCA")
                    {
                        Functions.Save(this.dgv, dir);
                    }
                    if (Functions.OpenV == true)
                    {
                        this.toolStripStatusLabel1.Text = " Cargado exitosamente el " + DateTime.Now.ToShortTimeString();
                        Info.Icon = SystemIcons.Information;
                        Info.BalloonTipText = "Guardando. Favor de esperar";
                        Info.ShowBalloonTip(2000);
                    }
                    else
                    {

                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                }
        }

        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!this.dgv.Rows[this.rowIndex].IsNewRow)
                {
                    this.dgv.Rows[0].Selected = true;
                    this.dgv.Rows.RemoveAt(this.rowIndex);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                    
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
                this.Close();
            }
        }

        private void loadnotes()
        {
            try
            {
                if (el1.Text == "GARDNER DENVER 2000")
                    Notes.Text = Properties.Settings.Default.GARDNERDENVER2000;
                if (el1.Text == "GARDNER DENVER 15W")
                    Notes.Text = Properties.Settings.Default.GARDNERDENVER15W;
                if (el1.Text == "MAYHEW 1000")
                    Notes.Text = Properties.Settings.Default.MAYHEW1000;
                if (el1.Text == "MIDWAY 15M")
                    Notes.Text = Properties.Settings.Default.MIDWAY15M;
                if (el1.Text == "IR CYCLON RO 300")
                    Notes.Text = Properties.Settings.Default.IRCYCLONRO300;
                if (el1.Text == "BUCYRUS ERIE 22")
                    Notes.Text = Properties.Settings.Default.BUCYRUSERIE22;
                if (el1.Text == "BUCYRUS ERIE 36L")
                    Notes.Text = Properties.Settings.Default.BUCYRUSERIE36L;
                if (el1.Text == "INGERSOLL RAND (1)")
                    Notes.Text = Properties.Settings.Default.INGERSOLLRAND1;
                if (el1.Text == "INGERSOLL RAND (2)")
                    Notes.Text = Properties.Settings.Default.INGERSOLLRAND2;
                if (el1.Text == "SULLAIR")
                    Notes.Text = Properties.Settings.Default.SULLAIR;
                if (el1.Text == "NCA")
                    Notes.Text = Properties.Settings.Default.NCA;
            }
            catch { }
            //Properties.Settings.Default.Save();
        }

        private void savenotes()
        {
            try
            {
                if (el1.Text == "GARDNER DENVER 2000")
                    Properties.Settings.Default.GARDNERDENVER2000 = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "GARDNER DENVER 15W")
                    Properties.Settings.Default.GARDNERDENVER15W = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "MAYHEW 1000")
                    Properties.Settings.Default.MAYHEW1000 = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "MIDWAY 15M")
                    Properties.Settings.Default.MIDWAY15M = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "IR CYCLON RO 300")
                    Properties.Settings.Default.IRCYCLONRO300 = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "BUCYRUS ERIE 22")
                    Properties.Settings.Default.BUCYRUSERIE22 = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "BUCYRUS ERIE 36L")
                    Properties.Settings.Default.BUCYRUSERIE36L = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "INGERSOLL RAND (1)")
                    Properties.Settings.Default.INGERSOLLRAND1 = Notes.Text;
                    Properties.Settings.Default.Save();
                if (el1.Text == "INGERSOLL RAND (2)")
                    Properties.Settings.Default.INGERSOLLRAND2 = Notes.Text;
                    Properties.Settings.Default.Save();
                if (el1.Text == "SULLAIR")
                    Properties.Settings.Default.SULLAIR = Notes.Text;
                Properties.Settings.Default.Save();
                if (el1.Text == "NCA")
                    Properties.Settings.Default.NCA = Notes.Text;
                Properties.Settings.Default.Save();
            }
            catch {
                MessageBox.Show("Error guardando notas, Error");
                  }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            DateTime dtCurrentTime = DateTime.Now;
            Time.Text = dtCurrentTime.ToLongTimeString();
        }

        private void dgv_DataError_1(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception != null)
            {
            }
        }

        private void dgv_CellEndEdit_1(object sender, DataGridViewCellEventArgs e)
        {
            Calculate();
        }

        private void dgv_CellMouseDown_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            this.rowIndex = e.RowIndex;
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    this.rowIndex = e.RowIndex;
                    this.MenuStrip.Show(Cursor.Position);
                }
            }
            catch
            {

            }
        }

        private void dgv_CellMouseLeave_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                this.dgv.EndEdit();
                Calculate();

            }
            catch
            {
            }
        }

        private void dgv_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgv.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            this.dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgv.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical;
            this.dgv.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.dgv.GridColor = System.Drawing.Color.Silver;
            Functions.Fill(dgv);
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            
        }
    }
}
