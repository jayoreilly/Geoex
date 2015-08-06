using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Reflection;
using Geoex.Properties;
using System.ComponentModel;
using System.Drawing;
using System.Security.Permissions;




namespace Geoex
{
    class Functions
    {
        #region Open
        public static string D;
        public static string T;
        public static bool OpenV;

        public static void Open(DataGridView dS, System.Windows.Forms.BindingSource bS, string savepath)
        {
            OpenFileDialog filedlgExcel = new OpenFileDialog();
            filedlgExcel.Title = "Seleccionar archivo";
            filedlgExcel.InitialDirectory = savepath;
            filedlgExcel.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            filedlgExcel.FilterIndex = 1;
            if (filedlgExcel.ShowDialog() == DialogResult.OK)
            {
                OpenV = true;
                String path = filedlgExcel.FileName;
                D = path;
                T = filedlgExcel.SafeFileName;
                String connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                OleDbConnection conn = new OleDbConnection(connStr);
                string strSQL = "SELECT * FROM [Hoja1$]";
                OleDbCommand cmd = new OleDbCommand(strSQL, conn);
                System.Data.DataTable dT = new System.Data.DataTable();
                conn.Open();
                OleDbDataReader dR = cmd.ExecuteReader();
                dT.Load(dR);
                bS.DataSource = dT;
                dS.DataSource = bS;
                conn.Close();
            }
            else
            {
                OpenV = false;
            }
        }
        #endregion
        #region OnFormLoad
        public static void OpenedOnFormLoad(string path, DataGridView dgv, BindingSource bS)
        {
            String connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(connStr);
            string strSQL = "SELECT * FROM [Hoja1$]";
            OleDbCommand cmd = new OleDbCommand(strSQL, conn);
            System.Data.DataTable dT = new System.Data.DataTable();
            conn.Open();
            try
            {
                    OleDbDataReader dR = cmd.ExecuteReader();
                    dT.Load(dR);
                    bS.DataSource = dT;
                    dgv.DataSource = bS;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        #endregion
        #region Save
        public static void Save(DataGridView dgv, string Savepath)
        {
            Microsoft.Office.Interop.Excel.Application APP = new Microsoft.Office.Interop.Excel.Application();
            APP.Application.Workbooks.Add(Type.Missing);
            APP.Columns.ColumnWidth = 20;
            for (int i = 1; i < dgv.Columns.Count + 1; i++)
            {
                APP.Cells[1, i] = dgv.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {

                    APP.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value;
                }
            }
            FileInfo file = new FileInfo(Savepath);
            if (file.Exists)
            {
                file.Delete();
            }
            OpenV = true;
            APP.ActiveWorkbook.SaveAs(Savepath.ToString());
            APP.ActiveWorkbook.Saved = true;
            APP.Quit();
        }
        #endregion
        #region Saveas
        public static void Saveas(DataGridView dgv, string path)
        {
            Microsoft.Office.Interop.Excel.Application APP = new Microsoft.Office.Interop.Excel.Application();
            APP.Application.Workbooks.Add(Type.Missing);
            APP.Columns.ColumnWidth = 20;
            for (int i = 1; i < dgv.Columns.Count + 1; i++)
            {
                APP.Cells[1, i] = dgv.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {

                    APP.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value;
                }
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel File (*.xlsx)|*.xlsx";
            saveFileDialog1.InitialDirectory = path;
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.OverwritePrompt = false;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(saveFileDialog1.FileName);
                if (file.Exists)
                {
                    file.Delete();
                }
                APP.ActiveWorkbook.SaveAs(saveFileDialog1.FileName.ToString());
                APP.ActiveWorkbook.Saved = true;
                APP.Quit();
                OpenV = true;
            }
        }
        #endregion
        #region SaveasClose
        public static void SaveasClose(DataGridView dgv, string Savepath)
        {
            Microsoft.Office.Interop.Excel.Application APP = new Microsoft.Office.Interop.Excel.Application();
            APP.Application.Workbooks.Add(Type.Missing);
            APP.Columns.ColumnWidth = 20;
            for (int i = 1; i < dgv.Columns.Count + 1; i++)
            {
                APP.Cells[1, i] = dgv.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {

                    APP.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value;
                }
            }
            FileInfo file = new FileInfo(Savepath);
            if (file.Exists)
            {
                file.Delete();
            }
            APP.ActiveWorkbook.SaveAs(Savepath.ToString());
            APP.ActiveWorkbook.Saved = true;
            APP.Quit();
        
        }
        #endregion
        #region Filldgv
        public static void Fill(DataGridView dgv)
        {
            try
            {
               
                    for (int j = 0; j < dgv.ColumnCount; j++)
                    {
                        dgv.Columns[j].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        dgv.Columns[j].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;


                    }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
        #region StringFileSafeName
        public static string SafeFileName()
        {
            return T;
        }

        #endregion
        #region newFile

        public static void NewFile(DataGridView dS, System.Windows.Forms.BindingSource bS, string dir)
        {

            OpenFileDialog filedlgExcel = new OpenFileDialog();
            filedlgExcel.Title = "Seleccionar archivo";
            filedlgExcel.InitialDirectory = dir;
            filedlgExcel.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            filedlgExcel.FilterIndex = 1;
            if (filedlgExcel.ShowDialog() == DialogResult.OK)
            {
                OpenV = true;
                String path = filedlgExcel.FileName;
                D = path;
                T = filedlgExcel.SafeFileName;
                String connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                OleDbConnection conn = new OleDbConnection(connStr);
                string strSQL = "SELECT * FROM [Hoja1$]";
                OleDbCommand cmd = new OleDbCommand(strSQL, conn);
                System.Data.DataTable dT = new System.Data.DataTable();
                conn.Open();
                OleDbDataReader dR = cmd.ExecuteReader();
                dT.Load(dR);
                bS.DataSource = dT;
                dS.DataSource = bS;
                conn.Close();
            }
            else
            {
                OpenV = false;
            }
        }


#endregion
        #region FadeInOut
        public static async void FadeIn(Form o,int interval = 80)
        {
            while (o.Opacity < 1.0)
            {
                await Task.Delay(interval);
                o.Opacity += 0.05;
            }
            o.Opacity = 1;
        }
    
        public static async void FadeOut(Form o,int interval = 80)
       {
          
            while (o.Opacity > 0.0)
            {
                await Task.Delay(interval);
                o.Opacity -= 0.05;
            }
            o.Opacity = 0;
          
        }


        #endregion
    }

}

