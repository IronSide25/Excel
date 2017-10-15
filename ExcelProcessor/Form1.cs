using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProcessor
{
    public partial class Form1 : Form
    {
        List<Truck> TruckData = new List<Truck>();
        List<String> Alphabet = new List<String>();
        Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();


        public Form1()
        {           
            Alphabet.Add("A"); Alphabet.Add("B"); Alphabet.Add("C"); Alphabet.Add("D"); Alphabet.Add("E"); Alphabet.Add("F"); Alphabet.Add("G"); Alphabet.Add("H"); Alphabet.Add("I"); Alphabet.Add("J"); Alphabet.Add("K"); Alphabet.Add("L"); Alphabet.Add("M"); Alphabet.Add("N"); Alphabet.Add("O"); Alphabet.Add("P"); Alphabet.Add("Q"); Alphabet.Add("R"); Alphabet.Add("S"); Alphabet.Add("T"); Alphabet.Add("U"); Alphabet.Add("V"); Alphabet.Add("W"); Alphabet.Add("X"); Alphabet.Add("Y"); Alphabet.Add("Z");

            


            string OutputPath = @"D:\Data\Programming\Visual C#\Excel\sample data\example-report\Output.txt";

            string xlsFileBelgien = @"D:\Data\Programming\Visual C#\Excel\sample data\truck-data\Belgien 2017 Monatsinfos .xls";
            string xlsFileEast = @"D:\Data\Programming\Visual C#\Excel\sample data\truck-data\EAST 2017 Monatsinfos .xls";
            string xlsFileWest = @"D:\Data\Programming\Visual C#\Excel\sample data\truck-data\WEST 2017 Monatsinfos .xls";
            string ExportGridData1 = @"D:\Data\Programming\Visual C#\Excel\sample data\example-report\ExportGridData20170824 15 41 33.xlsb";
            string ExportGridData2 = @"D:\Data\Programming\Visual C#\Excel\sample data\example-report\ExportGridData20170824 15 28 26.xlsb";
            string _300606Path = @"D:\Data\Programming\Visual C#\Excel\sample data\example-report\300606.XLSX";
            string _F61506817081Path = @"D:\Data\Programming\Visual C#\Excel\sample data\example-report\F61506817081.xlsx";
            
            //USUN WARTOSCI UJEMNE 

            InitializeComponent();
            TruckDataToExcel("4",xlsFileBelgien);
            TruckDataToExcel("4", xlsFileEast);
            TruckDataToExcel("4", xlsFileWest);
            ExportGridDataToExcel(ExportGridData1);
            ExportGridDataToExcel(ExportGridData2);
            _F61506817081ToExcel(_F61506817081Path);

            SaveOutputToTxt(OutputPath);

            MyExcel.Workbooks.Close();
            MyExcel.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void ExcelExample()
        {
            string xlFile = @"D:\Data\Programowanie\Visual C#\ExcelProcessor\sample data\truck-data\Belgien 2017 Monatsinfos .xls";

            //Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            MyExcel.Workbooks.Open(xlFile);
            MyExcel.Visible = true;
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;

            string CellItemValue = System.Convert.ToString(MyCells.Item[7, "C"].Value);

            MessageBox.Show(CellItemValue, "cell Value");

            //MyCells.Item[1, "A"].Value = "123";
        }

        private void TruckDataToExcel(String Month, string xlFile)
        {
            string xlFileBelgien = xlFile;
            
            //Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            MyExcel.Workbooks.Open(xlFileBelgien);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;

            int iRowCount = MyWorksheet.UsedRange.Rows.Count;
            int iColumnCount = MyWorksheet.UsedRange.Columns.Count;
            
            int KilometersStartRow;
            String KilometersStartColumn;

            //search for month and kilemeters column
            int SearchCurrentColumnCounter = 1;
            String CurrentColumn = "A";

            while (System.Convert.ToString(MyCells.Item[2, CurrentColumn].Value) != Month)
            {
                SearchCurrentColumnCounter++;
                if(SearchCurrentColumnCounter>25)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(Alphabet[(SearchCurrentColumnCounter / 26)-1]);
                    sb.Append(Alphabet[SearchCurrentColumnCounter % 26]);
                    CurrentColumn = sb.ToString();
                }
                else
                    CurrentColumn = Alphabet[SearchCurrentColumnCounter];
            }

            SearchCurrentColumnCounter++;
            if (SearchCurrentColumnCounter > 25)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(Alphabet[(SearchCurrentColumnCounter / 26) - 1]);
                sb.Append(Alphabet[SearchCurrentColumnCounter % 26]);
                CurrentColumn = sb.ToString();
            }
            else
                CurrentColumn = Alphabet[SearchCurrentColumnCounter];

            KilometersStartRow = 7;
            KilometersStartColumn = CurrentColumn;
            //searching done

            int RegistrationStartRow = 7;
            String RegistrationStartColumn = "E";

            while (MyCells.Item[RegistrationStartRow, RegistrationStartColumn].Value!=null)
            {
                String Reg = MyCells.Item[RegistrationStartRow, RegistrationStartColumn].Value;
                Reg = Reg.Replace(" ", string.Empty);
                
                TruckData.Add(new Truck { Registration = Reg, Kilometers = System.Convert.ToInt32(MyCells.Item[KilometersStartRow, KilometersStartColumn].Value) });
                RegistrationStartRow++;
                KilometersStartRow++;               
            }
           
        }        

        private void ExportGridDataToExcel(string Path)
        {
            string RegistrationColumn = "B";
            string ProductColumn = "G";
            string QuantityColumn = "H";
            string NettoPriceColumn = "M";
            string CurrencyColumn = "I";

            //Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;
            
            MyExcel.Workbooks.Open(Path);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;

            int CurrentRow = 2;
            int iRowCount = MyWorksheet.UsedRange.Rows.Count;

            while(CurrentRow<=iRowCount)
            {
                String Reg = MyCells.Item[CurrentRow, RegistrationColumn].Value;
                Reg = Reg.Replace(" ", string.Empty);

                foreach (Truck element in TruckData)
                {
                    if(element.Registration==Reg)
                    {
                        String ProductName = MyCells.Item[CurrentRow, ProductColumn].Value;
                        if (match(ProductName, new string[]{"Olej","Diesel"}))
                        {
                            element.DieselL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.DieselCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else if (match(ProductName, new string[] { "Autostrada", "Podatek", "Road tax", "Eurovignette", "Motorway" }))
                        {
                            element.RoadTax += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else if (match(ProductName, new string[] {"AdBlue" }))
                        {
                            element.AdblueL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.AdblueCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else  //OTHER COST TO MAJA BYC M.IN. NIEOPISANE??
                        {
                            element.OtherCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                    }                   
                }
                CurrentRow++;

               
            }

            
        }

        private void _300606ToExcel (string Path)
        {

        }

        private void _F61506817081ToExcel(string Path)
        {
            string RegistrationColumn = "X";
            string ProductColumn = "E";
            string QuantityColumn = "F";
            string NettoPriceColumn = "P";
            string CurrencyColumn = "O";

            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            MyExcel.Workbooks.Open(Path);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;

            int CurrentRow = 2;
            int iRowCount = MyWorksheet.UsedRange.Rows.Count;

            while (CurrentRow <= iRowCount)
            {
                String Reg = MyCells.Item[CurrentRow, RegistrationColumn].Value;
                Reg = Reg.Replace(" ", string.Empty);

                foreach (Truck element in TruckData)
                {
                    if (element.Registration == Reg)
                    {
                        String ProductName = MyCells.Item[CurrentRow, ProductColumn].Value;
                        if (match(ProductName, new string[] { "Olej", "Diesel" }))
                        {
                            element.DieselL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.DieselCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else if (match(ProductName, new string[] { "Autostrada", "Podatek", "Road tax", "Eurovignette", "Motorway" }))
                        {
                            element.RoadTax += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else if (match(ProductName, new string[] { "AdBlue" }))
                        {
                            element.AdblueL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.AdblueCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else  //OTHER COST TO MAJA BYC M.IN. NIEOPISANE??
                        {
                            element.OtherCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                    }
                }
                CurrentRow++;


            }

        }






        public bool match(String ProductName, string [] Tab)
        {
            bool check = false;

            foreach (string element in Tab)
            {
                if(ProductName.Contains(element))
                {
                    check = true;
                }
            }
            return check;
        }

        public void SaveOutputToTxt(string Path)
        {       
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(Path))
            {
                foreach (Truck element in TruckData)
                {                   
                    String R = element.GetRegistration();
                    file.WriteLine(R);
                    float K = element.GetKilometers();
                    file.WriteLine(K);
                    K = element.AdblueCost;
                    file.WriteLine(K);
                    K = element.AdblueL;
                    file.WriteLine(K);
                    K = element.DieselCost;
                    file.WriteLine(K);
                    K = element.DieselL;
                    file.WriteLine(K);
                    K = element.RoadTax;
                    file.WriteLine(K);
                    K = element.OtherCost;
                    file.WriteLine(K);

                    file.WriteLine(System.Environment.NewLine);
                }
            }
        }

        

    }
















    public class Truck
    {
        public String Registration;
        public int Kilometers;
        public float DieselL;
        public float DieselCost;
        public float AdblueL;
        public float AdblueCost;
        public float RoadTax;
        public float OtherCost;

        public Truck()
        {
        DieselL=0;
        DieselCost=0;
        AdblueL=0;
        AdblueCost=0;
        RoadTax=0;
        OtherCost=0;
        }

        public String GetRegistration()
        {
            return this.Registration;
        }

        public int GetKilometers()
        {
            return this.Kilometers;
        }

    }




}




