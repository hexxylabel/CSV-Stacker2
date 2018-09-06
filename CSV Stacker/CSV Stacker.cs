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
        public Form1()
        {
            InitializeComponent();            
        }

        private void button3_Click(object sender, EventArgs e)
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
                    var fileNameWFP = fileNameWN;
                    var poweradeTrue = false;
                    if (fileNameWFP[0] == 'P')
                    {
                        fileNameWFP = fileNameWFP.Substring(1, fileNameWFP.Length - 1);
                        poweradeTrue = true;
                    }
                    char fileNameLast = fileNameWFP[fileNameWFP.Length - 1];
                    char fileNameLeg = fileNameWFP[fileNameWFP.Length - 2];
                    //double age = Char.GetNumericValue(fileNameWE[0]) * 10 + Char.GetNumericValue(fileName[1]);

                    if (fileNameWFP == textBox1.Text) {
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
                            //////////////////////File name without extension///////////////////////////
                            if (d == 16)
                            {
                                oSheet.Cells[row, column] = fileNameWE;
                                column++;
                            }
                            /////////////////////////////////////////////////

                            /////////////////////Name////////////////////////////
                            if (d == 17 && fileNameLast != 'j' && fileNameLeg != 'l' && fileNameLast != 'l')
                            {
                                oSheet.Cells[row, column] = fileNameWFP;
                                column++;
                            }
                            else if (d == 17 && fileNameLast == 'j' && fileNameLeg != 'l')
                            {
                                oSheet.Cells[row, column] = fileNameWFP.Remove(fileNameWFP.Length - 1);
                                column++;
                            }
                            else if (d == 17 && fileNameLast != 'j' && fileNameLast == 'l') {
                                oSheet.Cells[row, column] = fileNameWFP.Remove(fileNameWFP.Length - 1);
                                column++;
                            }
                            else if (d == 17 && fileNameLast == 'j' && fileNameLeg == 'l')
                            {
                                oSheet.Cells[row, column] = fileNameWFP.Remove(fileNameWFP.Length - 2);
                                column++;
                            }
                            /////////////////////////////////////////////////

                            /////////////////////Ivás////////////////////////////
                            if (d == 18 && count == 1 && !poweradeTrue)
                            {
                                oSheet.Cells[row, column] = "Normál";
                                column++;
                            }
                            else if (d == 18 && count == 2 && !poweradeTrue)
                            {
                                oSheet.Cells[row, column] = "5 perc";
                                column++;
                            }
                            else if (d == 18 && count == 3 && !poweradeTrue)
                            {
                                oSheet.Cells[row, column] = "10 perc";
                                column++;
                            }
                            else if (d == 18 && count == 4 && !poweradeTrue)
                            {
                                oSheet.Cells[row, column] = "15 perc";
                                column++;
                            }
                            else if (d == 18 && count == 1 && poweradeTrue)
                            {
                                oSheet.Cells[row, column] = "1. ivás";
                                column++;
                            }
                            else if (d == 18 && count == 2 && poweradeTrue)
                            {
                                oSheet.Cells[row, column] = "2. ivás";
                                column++;
                            }
                            /////////////////////////////////////////////////

                            /////////////////////Kar////////////////////////////
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
                            /////////////////////////////////////////////////

                            /////////////////////Testrész////////////////////////////               
                            if (d == 20 /* && !poweradeTrue */ && fileNameLeg != 'l' && fileNameLast != 'l')
                            {
                                oSheet.Cells[row, column] = "kar";
                                column++;
                            }
                            else if (d == 20 /* && !poweradeTrue */ && (fileNameLeg == 'l' || fileNameLast == 'l'))
                            {
                                oSheet.Cells[row, column] = "láb";
                                column++;
                            }
                            /*
                            else if (d == 20 && poweradeTrue)
                            {
                                oSheet.Cells[row, column] = "Powerade";
                                column++;
                            }
                            */
                            /////////////////////////////////////////////////

                            /////////////////////Kor////////////////////////////            
                            //if (d == 20)
                            //{
                            //    oSheet.Cells[row, column] = age;
                            //    column++;
                            //}                        
                            /////////////////////////////////////////////////

                            /////////////////////Adatok////////////////////////////
                            if (d >= 22 && d <= 114)
                            {
                                var values = lines.Split(',');
                                oSheet.Cells[row, column] = values[1];
                                column++;
                            }
                            /////////////////////////////////////////////////

                            /////////////////////////////////////////////////
                            if (d == 115)
                            {
                                break;
                            }
                            /////////////////////////////////////////////////

                            d++;
                        }
                        row++;
                    }

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
                    var fileNameWFP = fileNameWN;
                    var poweradeTrue = false;
                    if (fileNameWFP[0] == 'P')
                    {
                        fileNameWFP = fileNameWFP.Substring(1, fileNameWFP.Length - 1);
                        poweradeTrue = true;
                    }
                    char fileNameLast = fileNameWFP[fileNameWFP.Length - 1];
                    char fileNameLeg = fileNameWFP[fileNameWFP.Length - 2];

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
                        //////////////////////File name without extension///////////////////////////
                        if (d == 16)
                        {
                            oSheet.Cells[row, column] = fileNameWE;
                            column++;
                        }
                        /////////////////////////////////////////////////

                        /////////////////////Name////////////////////////////
                        if (d == 17 && fileNameLast != 'j' && fileNameLeg != 'l' && fileNameLast != 'l')
                        {
                            oSheet.Cells[row, column] = fileNameWFP;
                            column++;
                        }
                        else if (d == 17 && fileNameLast == 'j' && fileNameLeg != 'l')
                        {
                            oSheet.Cells[row, column] = fileNameWFP.Remove(fileNameWFP.Length - 1);
                            column++;
                        }
                        else if (d == 17 && fileNameLast != 'j' && fileNameLast == 'l')
                        {
                            oSheet.Cells[row, column] = fileNameWFP.Remove(fileNameWFP.Length - 1);
                            column++;
                        }
                        else if (d == 17 && fileNameLast == 'j' && fileNameLeg == 'l')
                        {
                            oSheet.Cells[row, column] = fileNameWFP.Remove(fileNameWFP.Length - 2);
                            column++;
                        }
                        /////////////////////////////////////////////////

                        /////////////////////Ivás////////////////////////////
                        if (d == 18 && count == 1 && !poweradeTrue)
                        {
                            oSheet.Cells[row, column] = "Normál";
                            column++;
                        }
                        else if (d == 18 && count == 2 && !poweradeTrue)
                        {
                            oSheet.Cells[row, column] = "5 perc";
                            column++;
                        }
                        else if (d == 18 && count == 3 && !poweradeTrue)
                        {
                            oSheet.Cells[row, column] = "10 perc";
                            column++;
                        }
                        else if (d == 18 && count == 4 && !poweradeTrue)
                        {
                            oSheet.Cells[row, column] = "15 perc";
                            column++;
                        }
                        else if (d == 18 && count == 1 && poweradeTrue)
                        {
                            oSheet.Cells[row, column] = "1. ivás";
                            column++;
                        }
                        else if (d == 18 && count == 2 && poweradeTrue)
                        {
                            oSheet.Cells[row, column] = "2. ivás";
                            column++;
                        }
                        /////////////////////////////////////////////////

                        /////////////////////Kar////////////////////////////
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
                        /////////////////////////////////////////////////

                        /////////////////////Testrész////////////////////////////               
                        if (d == 20 /* && !poweradeTrue */ && fileNameLeg != 'l' && fileNameLast != 'l')
                        {
                            oSheet.Cells[row, column] = "kar";
                            column++;
                        }
                        else if (d == 20 /* && !poweradeTrue */ && (fileNameLeg == 'l' || fileNameLast == 'l'))
                        {
                            oSheet.Cells[row, column] = "láb";
                            column++;
                        }
                        /*
                        else if (d == 20 && poweradeTrue)
                        {
                            oSheet.Cells[row, column] = "Powerade";
                            column++;
                        }
                        */
                        /////////////////////////////////////////////////

                        /////////////////////Kor////////////////////////////            
                        //if (d == 20)
                        //{
                        //    oSheet.Cells[row, column] = age;
                        //    column++;
                        //}                        
                        /////////////////////////////////////////////////

                        /////////////////////Adatok////////////////////////////
                        if (d >= 22 && d <= 114)
                        {
                            var values = lines.Split(',');
                            oSheet.Cells[row, column] = values[1];
                            column++;
                        }
                        /////////////////////////////////////////////////

                        /////////////////////////////////////////////////
                        if (d == 115)
                        {
                            break;
                        }
                        /////////////////////////////////////////////////

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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
