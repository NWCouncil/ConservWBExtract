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
            // First search the files in the parent directory
            foreach (string strFile in Directory.GetFiles(textBox1.Text,"*.*", SearchOption.AllDirectories))
            {
                Match match = Regex.Match(strFile, @".*\.xls.*", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    // You've found an Excel file...
                    System.Diagnostics.Debug.WriteLine(strFile);
                }
            }
            this.Visible = true;
        }
    }
}
