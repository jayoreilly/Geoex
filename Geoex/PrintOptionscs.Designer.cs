namespace Geoex
{
    partial class PrintOptions
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PrintOptions));
            this.chklst = new System.Windows.Forms.CheckedListBox();
            this.metroButton1 = new MetroFramework.Controls.MetroButton();
            this.metroButton2 = new MetroFramework.Controls.MetroButton();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.chkFitToPageWidth = new MetroFramework.Controls.MetroCheckBox();
            this.rdoAllRows = new MetroFramework.Controls.MetroCheckBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // chklst
            // 
            this.chklst.CheckOnClick = true;
            this.chklst.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chklst.FormattingEnabled = true;
            this.chklst.Location = new System.Drawing.Point(10, 90);
            this.chklst.Name = "chklst";
            this.chklst.Size = new System.Drawing.Size(170, 244);
            this.chklst.TabIndex = 14;
            // 
            // metroButton1
            // 
            this.metroButton1.Location = new System.Drawing.Point(197, 311);
            this.metroButton1.Name = "metroButton1";
            this.metroButton1.Size = new System.Drawing.Size(75, 23);
            this.metroButton1.Style = MetroFramework.MetroColorStyle.Blue;
            this.metroButton1.TabIndex = 21;
            this.metroButton1.Text = "&Si";
            this.metroButton1.Theme = MetroFramework.MetroThemeStyle.Light;
            this.metroButton1.UseSelectable = true;
            this.metroButton1.Click += new System.EventHandler(this.metroButton1_Click);
            // 
            // metroButton2
            // 
            this.metroButton2.Location = new System.Drawing.Point(278, 311);
            this.metroButton2.Name = "metroButton2";
            this.metroButton2.Size = new System.Drawing.Size(75, 23);
            this.metroButton2.Style = MetroFramework.MetroColorStyle.Blue;
            this.metroButton2.TabIndex = 22;
            this.metroButton2.Text = "&Cancelar";
            this.metroButton2.Theme = MetroFramework.MetroThemeStyle.Light;
            this.metroButton2.UseSelectable = true;
            this.metroButton2.Click += new System.EventHandler(this.metroButton2_Click);
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.Location = new System.Drawing.Point(9, 66);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(133, 19);
            this.metroLabel1.Style = MetroFramework.MetroColorStyle.Blue;
            this.metroLabel1.TabIndex = 23;
            this.metroLabel1.Text = "Columnas a imprimir";
            this.metroLabel1.Theme = MetroFramework.MetroThemeStyle.Light;
            // 
            // chkFitToPageWidth
            // 
            this.chkFitToPageWidth.AutoSize = true;
            this.chkFitToPageWidth.Checked = true;
            this.chkFitToPageWidth.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFitToPageWidth.Location = new System.Drawing.Point(211, 257);
            this.chkFitToPageWidth.Name = "chkFitToPageWidth";
            this.chkFitToPageWidth.Size = new System.Drawing.Size(107, 15);
            this.chkFitToPageWidth.Style = MetroFramework.MetroColorStyle.Silver;
            this.chkFitToPageWidth.TabIndex = 29;
            this.chkFitToPageWidth.Text = "  Centrar pagina";
            this.chkFitToPageWidth.Theme = MetroFramework.MetroThemeStyle.Light;
            this.chkFitToPageWidth.UseSelectable = true;
            // 
            // rdoAllRows
            // 
            this.rdoAllRows.AutoSize = true;
            this.rdoAllRows.Checked = true;
            this.rdoAllRows.CheckState = System.Windows.Forms.CheckState.Checked;
            this.rdoAllRows.Location = new System.Drawing.Point(211, 204);
            this.rdoAllRows.Name = "rdoAllRows";
            this.rdoAllRows.Size = new System.Drawing.Size(124, 30);
            this.rdoAllRows.Style = MetroFramework.MetroColorStyle.Silver;
            this.rdoAllRows.TabIndex = 30;
            this.rdoAllRows.Text = "   Filas a imprimir\r\n \"Seleccionar todo\"";
            this.rdoAllRows.Theme = MetroFramework.MetroThemeStyle.Light;
            this.rdoAllRows.UseSelectable = true;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // PrintOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(374, 357);
            this.Controls.Add(this.rdoAllRows);
            this.Controls.Add(this.chkFitToPageWidth);
            this.Controls.Add(this.metroLabel1);
            this.Controls.Add(this.metroButton2);
            this.Controls.Add(this.metroButton1);
            this.Controls.Add(this.chklst);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PrintOptions";
            this.Opacity = 0.98D;
            this.Resizable = false;
            this.ShadowType = MetroFramework.Forms.MetroFormShadowType.None;
            this.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation;
            this.Style = MetroFramework.MetroColorStyle.Silver;
            this.Text = "Opciones de impresion";
            this.Load += new System.EventHandler(this.PrintOptionscs_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.CheckedListBox chklst;
        private MetroFramework.Controls.MetroButton metroButton1;
        private MetroFramework.Controls.MetroButton metroButton2;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private MetroFramework.Controls.MetroCheckBox chkFitToPageWidth;
        private MetroFramework.Controls.MetroCheckBox rdoAllRows;
        private System.Windows.Forms.Timer timer1;
    }
}