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
        public List<string> dirsList;


        public Form1()
        {
            InitializeComponent();

            //dirList initialization
            dirsList = new List<string>();

            //drag and drop actions
            this.AllowDrop = true;
            this.DragEnter += Form1_DragEnter;
            this.DragDrop += Form1_DragDrop;

            CurrencyConverter cur = new CurrencyConverter();
            
            //jebane gówno
            //MessageBox.Show(cur.getRateOf("EUR").ToString());

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
            filesList.SetItemChecked(filesList.Items.Count - 1, true);
        }

        private void filesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (filesList.SelectedIndex == -1)
                return;

            dirsList.RemoveAt(filesList.SelectedIndex);
            filesList.Items.Remove(filesList.SelectedItem);

        }

        void cancelGeneration(ExcelReader excelReader)
        {
            MessageBox.Show("Generowanie raportu przerwane.\n" +
                            "Sprawdź, czy wybrałeś właściwe pliki.\n" +
                            "Jeśli problem będzie występował dalej, zamknij wszystkie aktywne arkusze Excela");
            excelReader.closeAll();
            excelProcessingProgress.Value = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (month.SelectedIndex == -1)
            {
                MessageBox.Show("Wybierz miesiąc!");
                return;
            }

            DateTime today = DateTime.Today;
            string dateString = today.ToString("dd-MM-yyyy");

            string outputPath = "";
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "zestawienie_ciężarówek_" + dateString + ".xlsx";
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


            foreach (string path in dirsList)
            {
                ExcelType excelType = ExcelRecogniser.recognizeExcel(path);
                if (excelType == ExcelType.ERROR)
                {
                    cancelGeneration(excelReader);
                    return;
                }
                if (excelType == ExcelType.TRUCK_DATA)
                {
                    //MessageBox.Show("sending " + path + ", month: '" + selectedMonth + "'");
                    if (!excelReader.TruckDataToExcel(selectedMonth, path))
                    {
                        cancelGeneration(excelReader);

                        return;
                    }
                }if (excelType == ExcelType.EXTRA_INVOICE)
                {
                    if(!excelReader.extraInvoiceToExcel(path))
                    {
                        cancelGeneration(excelReader);
                    }


                }
                excelProcessingProgress.PerformStep();
            }

            /*
            //thats an old code, it's here for reasons
                    case ExcelType.EXTRA_INVOICE:
                        stepPerformedSuccessfully = excelReader.extraInvoiceToExcel(path);
                        break;
             */




            foreach (string path in dirsList)
            {
                ExcelType excelType = ExcelRecogniser.recognizeExcel(path);

                //MessageBox.Show("recognized: " + excelType.ToString());


                bool stepPerformedSuccessfully = true;

                switch (excelType)
                {
                    case ExcelType.EXPORT_GRID_DATA:

                        stepPerformedSuccessfully = excelReader.ExportGridDataToExcel(path);
                        break;

                    case ExcelType.F_AND_NUMBERS:
                        stepPerformedSuccessfully = excelReader._F61506817081ToExcel(path);
                        break;
                    case ExcelType.JUST_NUMBERS:
                        stepPerformedSuccessfully = excelReader._300606ToExcel(path);
                        break;
                    case ExcelType.SN_AND_NUMBERS:
                        stepPerformedSuccessfully = excelReader.SN760756ToExcel(path);
                        break;

                    case ExcelType.ERROR:
                        cancelGeneration(excelReader);
                        return;
                    

                    

                    default:
                        break;

                }

                if (!stepPerformedSuccessfully)
                {
                    cancelGeneration(excelReader);

                    return;
                }

                excelProcessingProgress.PerformStep();
                excelProcessingProgress.PerformStep();


            }

            try{
            if (File.Exists(outputPath))
                File.Delete(outputPath);
            

            excelReader.SaveOutputToFile(outputPath);
            excelProcessingProgress.PerformStep();

            excelReader.closeAll();
            excelProcessingProgress.PerformStep();

            MessageBox.Show("Zakończono!");

            status.Text = "Upuść pliki tutaj";
            excelProcessingProgress.Value = 0;
            Process.Start(outputPath);

            }catch (Exception exception){
                MessageBox.Show("Nie udało się zapisać pliku, ponieważ plik wybrany jako lokalizacja zapisu jest otwarty w innym programie");
                MessageBox.Show(exception.ToString());
                excelProcessingProgress.Value = 0;

            }


        }


        private void generate_extra_rachunek_Click(object sender, EventArgs e)
        {
            string outputPath = "";
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Gdzie zapisać przykładowy extra rachunek?";
            saveFileDialog1.FileName = "rachunek_extra.xlsx";
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

            //MessageBox.Show(outputPath);

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet = null;

            //ExcelApp.Visible = true;
            ExcelWorkBook = ExcelApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            //ExcelWorkBook.Worksheets.Add(); //Adding New Sheet in Excel Workbook
            
                ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; // Compulsory Line in which sheet you want to write data

                int CurrentRow = 2;
                int CurrentColumn = 1;

                ExcelWorkSheet.Cells[3, 1] = "Rejestracja";
                ExcelWorkSheet.Cells[3, 2] = "Koszt AdBlue";
                ExcelWorkSheet.Cells[3, 3] = "Diesel koszt";
                ExcelWorkSheet.Cells[3, 4] = "Ilosc AdBlue";
                ExcelWorkSheet.Cells[3, 5] = "Ilosc Diesel";
                ExcelWorkSheet.Cells[3, 6] = "Podatek Drogowy";
                ExcelWorkSheet.Cells[3, 7] = "Inne Koszty";
                ExcelWorkSheet.Cells[1, 1] = "Cena Diesel";
                ExcelWorkSheet.Cells[1, 2] = "Cena AdBlue";
                Microsoft.Office.Interop.Excel.Range aRange = ExcelWorkSheet.get_Range("A1", "Z6");
                aRange.Columns.AutoFit();
                ExcelWorkBook.Worksheets[1].Name = "Rachunek extra";

            if (File.Exists(outputPath))
                File.Delete(outputPath);

                ExcelWorkBook.SaveAs(outputPath);
                ExcelWorkBook.Close(false);
                ExcelApp.Quit();

            MessageBox.Show("Zapisano!");
            Process.Start(outputPath);

        }
    }
}
