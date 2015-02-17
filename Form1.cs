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
            var logfile = new StreamWriter(textBox3.Text);

            // Start Excel
            Excel.Application xlApp = new Excel.Application();

            // First search the files in the parent directory
            bool firstRPMWB = true;
            foreach (string strFile in Directory.GetFiles(textBox1.Text,"*.*", SearchOption.AllDirectories))
            {
                Match match = Regex.Match(strFile, @".*\.xls.*", RegexOptions.IgnoreCase);
                Match tempfile = Regex.Match(strFile, @"\~\$", RegexOptions.IgnoreCase);
                if (match.Success & !tempfile.Success)
                {
                    // You've found an Excel file...
                    Excel.Workbook xlWB;
                                      
                    xlWB = xlApp.Workbooks.Open(strFile, ReadOnly: true);
                    foreach (Excel._Worksheet xlWS in xlWB.Sheets)
                    {
                        System.Diagnostics.Debug.WriteLine(xlWS.Name);
                        if (xlWS.Name.ToUpper() == "FORRPM")
                        {
                            logfile.Write(strFile);
                            logfile.WriteLine();
                            String firstColValue = "NotEmpty";
                            Int16 i = 3;
                            while (firstColValue != null)
                            {
                                firstColValue = xlWS.Cells[i, 1].Value;
                                i++;
                            }
                            Excel.Range xlRng;
                            if (firstRPMWB)
                            {
                                xlRng = xlWS.get_Range((Excel.Range)xlWS.Cells[1, 1], (Excel.Range)xlWS.Cells[i - 2, 57]);
                                firstRPMWB = false;
                            }
                            else
                            {
                                xlRng = xlWS.get_Range((Excel.Range)xlWS.Cells[3, 1], (Excel.Range)xlWS.Cells[i - 2, 57]);
                            }
                            foreach (Excel.Range row in xlRng.Rows)
                            {
                                for (int j = 1; j < row.Columns.Count; j++)
                                {
                                    string writeVal = row.Cells[1, j].Text;
                                    double numVal;
                                    bool isNum = double.TryParse(writeVal, out numVal);
                                    if (row.Cells[1, j].Value2 != null)
                                    {
                                        if (!isNum)
                                        {
                                            outfile.Write("\"");
                                            outfile.Write(writeVal);
                                            outfile.Write("\"");
                                        }
                                        else
                                        {
                                            outfile.Write(row.Cells[1, j].Value2);
                                        }
                                    }
                                    else
                                    {
                                        outfile.Write("NA");
                                    }
                                    outfile.Write(",");
                                }
                                outfile.WriteLine();
                            }
                        }
                    }
                    xlWB.Close(SaveChanges: false);
                    outfile.Flush();
                    logfile.Flush();
                }
            }
            outfile.Close();
            logfile.Close();
            xlApp.Quit();
            this.Visible = true;
        }
        private void textBox1_Validating(object sender, CancelEventArgs e)
        {
            if (textBox1.Text != "")
            {
                Properties.Settings.Default["LastFolder"] = textBox1.Text;
            }
        }

        private void textBox2_Validating(object sender, CancelEventArgs e)
        {
            Match match = Regex.Match(textBox2.Text, @".*\.csv$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                textBox2.Text = textBox2.Text + ".csv";
            }
            if (textBox3.Text == "")
            {
                textBox3.Text = textBox2.Text.Replace(".csv", ".log");
            }
            if (textBox2.Text != "")
            {
                Properties.Settings.Default["LastFile"] = textBox2.Text;
            }
        }

        private void textBox3_Validating(object sender, CancelEventArgs e)
        {
            if (textBox3.Text != "")
            {
                Properties.Settings.Default["LastLog"] = textBox3.Text;
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
        }
    }
}
