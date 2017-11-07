using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;


//my cool stuff

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        //list of directories
        public List <string> dirsList;
       

        public Form1()
        {
            InitializeComponent();

            //dirList initialization
            dirsList = new List<string>();

            //drag and drop actions
            this.AllowDrop = true;
            this.DragEnter += Form1_DragEnter;
            this.DragDrop += Form1_DragDrop;

        }

        private void drop(object sender, DragEventArgs e)
        {
            MessageBox.Show("Test");
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
                foreach (string fileLoc in filePaths)
                {
                    // Code to read the contents of the text file
                    if (File.Exists(fileLoc))
                    {

                        addFile(fileLoc);
                        
                    }

                }
            }
        }

        private void addFile(string fileLoc)
        {
            //FIXME
            dirsList.Add(fileLoc);
            filesList.Items.Add(Path.GetFileName(fileLoc));
            filesList.SetItemChecked(filesList.Items.Count - 1, true) ;
        }

        private void filesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (filesList.SelectedIndex == -1)
                return;

            dirsList.RemoveAt(filesList.SelectedIndex);
            filesList.Items.Remove(filesList.SelectedItem);

        }
        

        private void button1_Click(object sender, EventArgs e)
        {

            if (month.SelectedIndex == -1)
            {
                MessageBox.Show("Wybierz miesiąc!");
                return;
            }

            string outputPath = "";
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "zestawienie_ciężarówek.xlsx";
            saveFileDialog1.Filter = "Arkusz Programu Microsoft Excel (*.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                outputPath = saveFileDialog1.FileName;
            }
            else
            {
                return;
            }

            status.Text = "Przetwarzanie..";
            string selectedMonth = (month.SelectedIndex + 1).ToString();


            excelProcessingProgress.Maximum = dirsList.Count * 3 + 2;
            excelProcessingProgress.Step = 1;


            ExcelReader excelReader = new ExcelReader();
            

            foreach(string path in dirsList)
            {
                if(ExcelRecogniser.recognizeExcel(path) == ExcelType.TRUCK_DATA)
                {
                    //MessageBox.Show("sending " + path + ", month: '" + selectedMonth + "'");
                    excelReader.TruckDataToExcel("4", path);
                }
                excelProcessingProgress.PerformStep();
            }


            

            foreach (string path in dirsList)
            {
                ExcelType excelType = ExcelRecogniser.recognizeExcel(path);

                //MessageBox.Show("recognized: " + excelType.ToString());

                switch (excelType)

                {
                    case ExcelType.EXPORT_GRID_DATA:

                        excelReader.ExportGridDataToExcel(path);
                        break;

                    case ExcelType.F_AND_NUMBERS:
                        excelReader._F61506817081ToExcel(path);
                        break;
                    case ExcelType.JUST_NUMBERS:
                        excelReader._300606ToExcel(path);
                        break;
                    case ExcelType.SN_AND_NUMBERS:
                        excelReader.SN760756ToExcel(path);
                        break;

                    default:
                        break;

                }

                excelProcessingProgress.PerformStep();
                excelProcessingProgress.PerformStep();


            }

            
            excelReader.SaveOutputToFile(outputPath);
            excelProcessingProgress.PerformStep();

            excelReader.closeAll();
            excelProcessingProgress.PerformStep();

            MessageBox.Show("Zakończono!");

            status.Text = "Upuść pliki tutaj";
            excelProcessingProgress.Value = 0;
            Process.Start(outputPath);


        }

        private void domainUpDown1_SelectedItemChanged(object sender, EventArgs e)
        {

        }


    }
}
