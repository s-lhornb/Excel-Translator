using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class ExcelForm : Form
    {
        private FolderBrowserDialog fbd = new FolderBrowserDialog();
        private FolderBrowserDialog fbd2 = new FolderBrowserDialog();
        private ExcelTranslator eT = new ExcelTranslator();
        private string[] filenames;

        public ExcelForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// On button press this method runs the translation process and schanges the progress bar and lable. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (Directory.Exists(PathBox.Text))
            {
                if (Directory.Exists(PathBox2.Text))
                {
                    filenames = eT.getFileList(PathBox.Text);
                    pBar.Minimum = 1;
                    pBar.Maximum = filenames.Length + 3;
                    pBar.Value = 1;

                    foreach (string file in filenames)
                    {
                        pLable.Text = file + " started to translate.";
                        eT.readExcel(file);
                        pBar.Value ++;
                    }

                    pLable.Text = "Creating File with issues.";
                    eT.createIssueReport(PathBox2.Text);
                    pBar.Value++;

                    pLable.Text = "Creating Excel-File.";
                    eT.createExcel(PathBox2.Text);
                    pBar.Value++;

                    eT.clean();

                    pLable.Text = "Done";
                }
                else
                {
                    MessageBox.Show("Invalid address for the output path");
                }
            }
            else
            {
                MessageBox.Show("Invalid address for the input path");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// This button opens the selection dialoge for the input path
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                PathBox.Text = fbd.SelectedPath;
        }

        /// <summary>
        /// This button opens the selection dialoge for the output path
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_2(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                PathBox2.Text = fbd.SelectedPath;
        }

        /// <summary>
        /// if the programm closes the excel translator is cleaned up
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            eT.clean();
        }
    }
}
