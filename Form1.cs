using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConservWBExtract
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = false;
            if (textBox1.Text != "")
            {
                folderDlg.SelectedPath = textBox1.Text;
            }
            // Show the file dialog
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK) {
                textBox1.Text = folderDlg.SelectedPath;
            }
        }

        private void btnFileName_Click(object sender, EventArgs e)
        {
            SaveFileDialog outputDlg = new SaveFileDialog();
            outputDlg.DefaultExt = ".csv";
            outputDlg.Filter = "Comma Separated Variable (*.csv)|*.csv";
            outputDlg.SupportMultiDottedExtensions = false;
            outputDlg.AddExtension = false;
            if (textBox2.Text != "")
            {
                outputDlg.FileName = textBox2.Text;
            }
            DialogResult result = outputDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox2.Text = outputDlg.FileName;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            // If no folder is selected then throw an error
            if (textBox1.Text == "")
            {
                MessageBox.Show("No Folder Selected", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // Open up the CSV file for output
            if (textBox2.Text == "")
            {
               MessageBox.Show("No Output File Selected", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            var outfile = new StreamWriter(textBox2.Text);

            // Start Excel
            Excel.Application xlApp = new Excel.Application();

            // First search the files in the parent directory
            foreach (string strFile in Directory.GetFiles(textBox1.Text,"*.*", SearchOption.AllDirectories))
            {
                Match match = Regex.Match(strFile, @".*\.xls.*", RegexOptions.IgnoreCase);
                Match tempfile = Regex.Match(strFile, @"\~\$", RegexOptions.IgnoreCase);
                if (match.Success & !tempfile.Success)
                {
                    // You've found an Excel file...
                    Excel.Workbook xlWB;
                                      
                    xlWB = xlApp.Workbooks.Open(strFile);
                    foreach (Excel._Worksheet xlWS in xlWB.Sheets)
                    {
                        System.Diagnostics.Debug.WriteLine(xlWS.Name);
                        if (xlWS.Name.ToUpper() == "FORRPM")
                        {
                            String firstColValue = "NotEmpty";
                            Int16 i = 3;
                            while (firstColValue != null)
                            {
                                firstColValue = xlWS.Cells[i, 1].Value;
                                i++;
                            }
                            Excel.Range xlRng = xlWS.get_Range((Excel.Range)xlWS.Cells[3, 1], (Excel.Range)xlWS.Cells[i-2, 53]);
                            foreach (Excel.Range row in xlRng.Rows)
                            {
                                for (int j = 1; j < row.Columns.Count; j++)
                                {
                                    //System.Diagnostics.Debug.Write(row.Cells[1, j].Value2);
                                    if (row.Cells[1, j].Value2 != null)
                                    {
                                        outfile.Write(row.Cells[1, j].Value2);
                                    }
                                    else
                                    {
                                        outfile.Write("NA");
                                    }
                                    outfile.Write(", ");
                                }
                                outfile.WriteLine();
                            }
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine(strFile);
                }
            }
            outfile.Close();
            xlApp.Quit();
            this.Visible = true;
        }
    }
}
