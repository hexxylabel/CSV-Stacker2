using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

using Excel = Microsoft.Office.Interop.Excel;

namespace CSV_Stacker
{
    public partial class Form1 : Form
    {
        public class nameAge
        {
            string[] name;
            int age;
        }

        public Form1()
        {
            InitializeComponent();            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] filePath;
            FolderBrowserDialog getFilePath;

            getFilePath = new FolderBrowserDialog();//File Browser open            
            getFilePath.Description = "Choose the data directory!";//File Broser Description

            if (getFilePath.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                label1.Text = "Selected folder: " + getFilePath.SelectedPath;//label text the selected folder
                Excel.Application oApp = new Excel.Application();// Open Excel app
                Excel.Workbook oBook = oApp.Workbooks.Add();// Create Excel Workbook
                Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);//Create Excel Worksheet

                progressBar1.Value = 0;//Progressbar set default
                filePath = Directory.GetFiles(getFilePath.SelectedPath, "*.csv", SearchOption.AllDirectories);//Get files

                progressBar1.Maximum = filePath.Length;//Set progressbar length
                int row = 1;

                var sorted = filePath.OrderBy(f => new FileInfo(f).Name.Count(char.IsDigit));//Sort the files
                
                foreach (string file in sorted)
                {
                    var fileName = Path.GetFileName(file);
                    var fileNameWE = Path.GetFileNameWithoutExtension(file);
                    var fileNameWN = Regex.Replace(fileNameWE, @"[\d]", string.Empty);
                    var fileNameWFN = fileNameWE.Substring(2, fileNameWE.Length - 2);

                    char fileNameLast = fileNameWN[fileNameWN.Length - 1];

                    //double age = Char.GetNumericValue(fileNameWE[0]) * 10 + Char.GetNumericValue(fileName[1]);

                    int count = fileName.Count(Char.IsDigit);

                    label3.Text = "Current file: " + fileName;
                    progressBar1.Increment(1);

                    int column = 1;
                    int d = 0;
                    string[] raw_line = File.ReadAllLines(file);

                    progressBar2.Value = 0;
                    progressBar2.Maximum = 115;
                    foreach (string lines in raw_line)
                    {
                        progressBar2.Increment(1);

                        if (d == 16)
                        {
                            oSheet.Cells[row, column] = fileNameWFN;
                            column++;
                        }
                        if (d == 17 && fileNameLast != 'j')
                        {
                            oSheet.Cells[row, column] = fileNameWN;
                            column++;
                        }
                        else if (d == 17 && fileNameLast == 'j')
                        {
                            oSheet.Cells[row, column] = fileNameWN.Remove(fileNameWN.Length - 1);
                            column++;
                        }
                        if (d == 18 && count == 3)
                        {
                            oSheet.Cells[row, column] = "Szomjas";
                            column++;
                        }
                        else if (d == 18 && count == 4)
                        {
                            oSheet.Cells[row, column] = "1";
                            column++;
                        }
                        else if (d == 18 && count == 5)
                        {
                            oSheet.Cells[row, column] = "2";
                            column++;
                        }
                        else if (d == 18 && count == 6)
                        {
                            oSheet.Cells[row, column] = "3";
                            column++;
                        }
                        if (d == 19 && fileNameLast != 'j')
                        {
                            oSheet.Cells[row, column] = "bal";
                            column++;
                        }
                        else if (d == 19 && fileNameLast == 'j')
                        {
                            oSheet.Cells[row, column] = "jobb";
                            column++;
                        }
                        //if (d == 20)
                        //{
                        //    oSheet.Cells[row, column] = age;
                        //    column++;
                        //}
                        if (d >= 22 && d <= 114)
                        {
                            var values = lines.Split(',');
                            oSheet.Cells[row, column] = values[1];
                            column++;
                        }
                        if (d == 115)
                        {
                            break;
                        }
                        d++;
                    }
                    row++;
                }

                var saveFileDialoge = new SaveFileDialog();
                saveFileDialoge.FileName = "output";
                saveFileDialoge.DefaultExt = "xlsx";
                saveFileDialoge.InitialDirectory = getFilePath.SelectedPath;

                if (saveFileDialoge.ShowDialog() == DialogResult.OK)
                {
                    oBook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    oBook.Close();
                    oApp.Quit();
                }
                else if (saveFileDialoge.ShowDialog() == DialogResult.Cancel || saveFileDialoge.ShowDialog() == DialogResult.Abort)
                {
                    oBook.Close(false);
                    oApp.Quit();
                }
            }            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void progressBar2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string[] filePath;
            FolderBrowserDialog getFilePath;

            getFilePath = new FolderBrowserDialog();//File Browser open            
            getFilePath.Description = "Choose the data directory!";//File Broser Description

            if (getFilePath.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                filePath = Directory.GetFiles(getFilePath.SelectedPath, "*.csv", SearchOption.AllDirectories);//Get files

                List < string > name = new List<string>();

                foreach (string file in filePath) {
                    var fileName = Path.GetFileName(file);
                    var fileNameWE = Path.GetFileNameWithoutExtension(file);
                    var fileNameWN = Regex.Replace(fileNameWE, @"[\d]", string.Empty);

                    name.Add(fileNameWN);
                    double age = Char.GetNumericValue(fileNameWE[0]) * 10 + Char.GetNumericValue(fileName[1]);                    
                    
                }

                var nameUnique = name.OrderBy(f => f).Distinct();
                

                foreach (string names in nameUnique)
                {
                    
                }
            }
        }
    }
}
