using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FileHelpers;

namespace WindowsFormsApp1
{
    [IgnoreFirst(5)]
    [IgnoreEmptyLines(true)]
    [DelimitedRecord(";")]
    public class TruckCSV
    {
        public string a, registration, c, d, e, f, g, h, i, j, k;
        public string product;

        [FieldConverter(ConverterKind.Double, ",")]
        public double quantity; //quantity
        public string n, o, p, q, r, s, t, u, v, w, x, y;

        [FieldConverter(ConverterKind.Double, ",")]
        public double netto_price;
        public string currency, ab, ac, ad, ae;
    }


    // RegistrationColumn = "B";
    // ProductColumn = "L";
    // QuantityColumn = "M";
    // NettoPriceColumn = "Z";
    // CurrencyColumn = "AA";



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
            DieselL = 0;
            DieselCost = 0;
            AdblueL = 0;
            AdblueCost = 0;
            RoadTax = 0;
            OtherCost = 0;
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

    public class ExcelReader
    {
        List<Truck> TruckData = new List<Truck>();
        List<String> Alphabet = new List<String>();
        Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
        CurrencyConverter currencyConverter;

        public ExcelReader()
        {
            currencyConverter = new CurrencyConverter();
            if(currencyConverter.hasBeenConstructedInAProperWay() == false)
            {
                System.Windows.Forms.MessageBox.Show("Nie działa internet!");
            }
            Alphabet.Add("A"); Alphabet.Add("B"); Alphabet.Add("C"); Alphabet.Add("D"); Alphabet.Add("E"); Alphabet.Add("F"); Alphabet.Add("G"); Alphabet.Add("H"); Alphabet.Add("I"); Alphabet.Add("J"); Alphabet.Add("K"); Alphabet.Add("L"); Alphabet.Add("M"); Alphabet.Add("N"); Alphabet.Add("O"); Alphabet.Add("P"); Alphabet.Add("Q"); Alphabet.Add("R"); Alphabet.Add("S"); Alphabet.Add("T"); Alphabet.Add("U"); Alphabet.Add("V"); Alphabet.Add("W"); Alphabet.Add("X"); Alphabet.Add("Y"); Alphabet.Add("Z");

        }


        public bool TruckDataToExcel(String Month, string xlFile)
        {
            string xlFileBelgien = xlFile;

            //Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            try
            {
                MyExcel.Workbooks.Open(xlFileBelgien);
                MyWorksheet = MyExcel.Worksheets.Item[1];
                MyCells = MyWorksheet.Cells;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + xlFile);
                return false;
            }

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
                if (SearchCurrentColumnCounter > 25)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(Alphabet[(SearchCurrentColumnCounter / 26) - 1]);
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

            while (MyCells.Item[RegistrationStartRow, RegistrationStartColumn].Value != null)
            {
                String Reg = MyCells.Item[RegistrationStartRow, RegistrationStartColumn].Value;
                Reg = Reg.Replace(" ", string.Empty);

                TruckData.Add(new Truck { Registration = Reg, Kilometers = System.Convert.ToInt32(MyCells.Item[KilometersStartRow, KilometersStartColumn].Value) });
                RegistrationStartRow++;
                KilometersStartRow++;
            }

            return true;

        }

        public bool extraInvoiceToExcel(string Path)
        {
            //Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            try
            {
                MyExcel.Workbooks.Open(Path);
                MyWorksheet = MyExcel.Worksheets.Item[1];
                MyCells = MyWorksheet.Cells;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + Path);
                return false;
            }

            /*
            * A - rejestracja
            * B - koszt adblue
            * C - koszt diesel
            * D - ilość adblue
            * E - ilość diesel
            * F - podatek drogowy
            * G - inne koszty
           */

            int RegistrationColumn = 1;
            int adBlueCostColumn = 2;
            int dieselCostColumn = 3;
            int adBlueAmountColumn = 4;
            int dieselAmountColumn = 5;
            int roadTaxColumn = 6;
            int otherCostsColumn = 7;

            int CurrentRow = 2;
            int iRowCount = MyWorksheet.UsedRange.Rows.Count;

            while (CurrentRow <= iRowCount)
            {


                String Reg = MyCells.Item[CurrentRow, RegistrationColumn].Value;
                Reg = Reg.Replace(" ", string.Empty);

               // System.Windows.Forms.MessageBox.Show(Reg);


                foreach (Truck element in TruckData)
                {
                    if (element.Registration == Reg) //this is working, just lookin bad
                    {
                        element.AdblueCost += MyCells.Item[CurrentRow, adBlueCostColumn].Value;
                        element.DieselCost += MyCells.Item[CurrentRow, dieselCostColumn].Value;

                        element.AdblueL += MyCells.Item[CurrentRow, adBlueAmountColumn].Value;
                        element.DieselL += MyCells.Item[CurrentRow, dieselAmountColumn].Value;

                        element.RoadTax += MyCells.Item[CurrentRow, roadTaxColumn].Value;
                        element.OtherCost += MyCells.Item[CurrentRow, otherCostsColumn].Value;
                    }
                }
                CurrentRow++;


            }

            return true;
        }

        public bool ExportGridDataToExcel(string Path)
        {
            string RegistrationColumn = "B";
            string ProductColumn = "G";
            string QuantityColumn = "H";
            string NettoPriceColumn = "M";
            string CurrencyColumn = "I";

            //Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            try { 
            MyExcel.Workbooks.Open(Path);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;
            }catch(Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + Path);
                return false;
            }

            int CurrentRow = 2;
            int iRowCount = MyWorksheet.UsedRange.Rows.Count;

            while (CurrentRow <= iRowCount)
            {
                String Reg = MyCells.Item[CurrentRow, RegistrationColumn].Value;
                Reg = Reg.Replace(" ", string.Empty);

                foreach (Truck element in TruckData)
                {
                    if (element.Registration == Reg)//this is working, just lookin bad
                    {
                        String ProductName = MyCells.Item[CurrentRow, ProductColumn].Value;
                        if (match(ProductName, new string[] { "Olej", "Diesel" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >= 0 && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                        {
                            element.DieselL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.DieselCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                        else if (match(ProductName, new string[] { "Autostrada", "Podatek", "Road tax", "Eurovignette", "Motorway" }) && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                        {
                            element.RoadTax += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                        else if (match(ProductName, new string[] { "AdBlue" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >= 0 && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                        {
                            element.AdblueL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.AdblueCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                        else if(MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0) 
                        {
                            element.OtherCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                    }
                }
                CurrentRow++;


            }

            return true;
        }


        public bool _300606ToExcel(string Path)
        {
            //return;
            //System.Windows.Forms.MessageBox.Show("siema");
            string RegistrationColumn = "N";
            string ProductColumn = "E";
            string QuantityColumn = "F";
            //string NettoPriceColumn = "P";
            //string CurrencyColumn = "O";

            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            try {
            MyExcel.Workbooks.Open(Path);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;
            } catch(Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + Path);
                return false;
            }

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
                        if (match(ProductName, new string[] { "1", "2" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >=0)
                        {
                            element.DieselL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            //element.DieselCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else if (match(ProductName, new string[] { "Autostrada", "Podatek", "Road tax", "Eurovignette", "Motorway" }))
                        {
                            //element.RoadTax += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else if (match(ProductName, new string[] { "5" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >= 0)
                        {
                            element.AdblueL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            //element.AdblueCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                        else  //OTHER COST TO MAJA BYC M.IN. NIEOPISANE??
                        {
                            //element.OtherCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value;
                            //currency
                        }
                    }
                }
                CurrentRow++;
            }

            return true;
        }

        public bool _F61506817081ToExcel(string Path)
        {
            string RegistrationColumn = "X";
            string ProductColumn = "E";
            string QuantityColumn = "F";
            string NettoPriceColumn = "P";
            string CurrencyColumn = "O";

            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            try {
            MyExcel.Workbooks.Open(Path);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;
            } catch(Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + Path);
                return false;
            }

            int CurrentRow = 2;
            int iRowCount = MyWorksheet.UsedRange.Rows.Count;

            while (CurrentRow <= iRowCount)
            {
                String Reg = MyCells.Item[CurrentRow, RegistrationColumn].Value;
                Reg = Reg.Replace(" ", string.Empty);

                foreach (Truck element in TruckData)
                {
                    if (element.Registration == Reg)//this is working, just lookin bad
                    {
                        String ProductName = MyCells.Item[CurrentRow, ProductColumn].Value;
                        if (match(ProductName, new string[] { "Olej", "Diesel" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >= 0 && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                        {
                            element.DieselL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.DieselCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                        else if (match(ProductName, new string[] { "Autostrada", "Podatek", "Road tax", "Eurovignette", "Motorway" }) && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                        {
                            element.RoadTax += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                        else if (match(ProductName, new string[] { "AdBlue" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >= 0 && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                        {
                            element.AdblueL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                            element.AdblueCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                        else if(MyCells.Item[CurrentRow, NettoPriceColumn].Value>=0) 
                        {
                            element.OtherCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                        }
                    }
                }
                CurrentRow++;
            }

            return true;
        }

        public bool SN760756ToExcel(string filePath)
        {

            //try {

            var engine = new FileHelperEngine<TruckCSV>();
            var result = engine.ReadFile(filePath);

            // result is now an array of RekordCSV

            foreach (TruckCSV csvTruck in result)
            {
                csvTruck.registration = csvTruck.registration.Replace(" ", string.Empty);


                //THIS IS KAROL'S CODE

                foreach (Truck element in TruckData)
                {
                    if (element.Registration == csvTruck.registration)//this is working, just lookin bad
                    {
                        //this is so fucking bad i don't even
                        String ProductName = csvTruck.product;

                        if (match(ProductName, new string[] { "Olej", "Diesel", "ON" }) && csvTruck.quantity >= 0 && csvTruck.netto_price >= 0)
                        {
                            element.DieselL += (float)csvTruck.quantity;
                            element.DieselCost += (float)csvTruck.netto_price * currencyConverter.getRateOf(csvTruck.currency);
                        }
                        else if (match(ProductName, new string[] { "Autostrada", "Podatek", "Road tax", "Eurovignette", "Motorway", "Eurowinieta", "drogowe" }) && csvTruck.netto_price >= 0)
                        {
                            element.RoadTax += (float)csvTruck.netto_price * currencyConverter.getRateOf(csvTruck.currency);
                        }
                        else if (match(ProductName, new string[] { "AdBlue" }) && csvTruck.quantity >= 0 && csvTruck.netto_price >= 0)
                        {
                            element.AdblueL += (float)csvTruck.quantity;
                            element.AdblueCost += (float)csvTruck.netto_price * currencyConverter.getRateOf(csvTruck.currency);
                        }
                        else if (csvTruck.netto_price >= 0)
                        {
                            element.OtherCost += (float)csvTruck.netto_price * currencyConverter.getRateOf(csvTruck.currency);
                        }
                    }
                }
                //END OF PIETRZAKOWY KOD

            }

        //}
                //catch (Exception e) {
                 //   System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + filePath);
                  //  return false;
               // }

            return true;
        }


        //mothafuckin' deprecated
        public bool SN760756ToExcel_xlsx(string Path)
        {
            //FIXME
            //return;

            string RegistrationColumn = "B";
            string ProductColumn = "L";
            string QuantityColumn = "M";
            string NettoPriceColumn = "Z";
            string CurrencyColumn = "AA";


            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;

            try
            {
                MyExcel.Workbooks.Open(Path);
                MyWorksheet = MyExcel.Worksheets.Item[1];
                MyCells = MyWorksheet.Cells;

            } catch(Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + Path);
                return false;
            }

            int CurrentRow = 6;
            int iRowCount = MyWorksheet.UsedRange.Rows.Count;

            while (CurrentRow <= iRowCount)
            {
                if (MyCells.Item[CurrentRow, RegistrationColumn].Value != null)
                {
                    String Reg = MyCells.Item[CurrentRow, RegistrationColumn].Value;
                    Reg = Reg.Replace(" ", string.Empty);

                    foreach (Truck element in TruckData)
                    {
                        if (element.Registration == Reg)//this is working, just lookin bad
                        {
                            String ProductName = MyCells.Item[CurrentRow, ProductColumn].Value;
                            if (match(ProductName, new string[] { "Olej", "Diesel", "ON" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >= 0 && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                            {
                                element.DieselL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                                element.DieselCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value*currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                            }
                            else if (match(ProductName, new string[] { "Autostrada", "Podatek", "Road tax", "Eurovignette", "Motorway", "Eurowinieta", "drogowe" }) && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                            {
                                element.RoadTax += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                            }
                            else if (match(ProductName, new string[] { "AdBlue" }) && MyCells.Item[CurrentRow, QuantityColumn].Value >= 0 && MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0)
                            {
                                element.AdblueL += MyCells.Item[CurrentRow, QuantityColumn].Value;
                                element.AdblueCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                            }
                            else if(MyCells.Item[CurrentRow, NettoPriceColumn].Value >= 0) 
                            {
                                element.OtherCost += MyCells.Item[CurrentRow, NettoPriceColumn].Value * currencyConverter.getRateOf(MyCells.Item[CurrentRow, CurrencyColumn].Value);
                            }
                        }
                    }
                }
                CurrentRow++;
            }

            return true;
        }




        public bool match(String ProductName, string[] Tab)
        {
            bool check = false;
            foreach (string element in Tab)
            {
                if (ProductName.ToLower().Contains(element.ToLower()))
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

        public void SaveOutputToFile(string Path)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet = null;

            //ExcelApp.Visible = true;
            ExcelWorkBook = ExcelApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            //ExcelWorkBook.Worksheets.Add(); //Adding New Sheet in Excel Workbook

            //try
            {
                ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; // Compulsory Line in which sheet you want to write data

                int CurrentRow = 2;
                int CurrentColumn = 1;

                ExcelWorkSheet.Cells[1, 1] = "Rejestracja";
                ExcelWorkSheet.Cells[1, 2] = "Kilometry";
                ExcelWorkSheet.Cells[1, 3] = "Koszt AdBlue";
                ExcelWorkSheet.Cells[1, 4] = "Ilosc AdBlue";
                ExcelWorkSheet.Cells[1, 5] = "Diesel koszt";
                ExcelWorkSheet.Cells[1, 6] = "Ilosc Diesel";
                ExcelWorkSheet.Cells[1, 7] = "Podatek Drogowy";
                ExcelWorkSheet.Cells[1, 8] = "Inne Koszty";


                ExcelWorkSheet.Cells[1, 9] = "Litry Diesel / 100 km";
                ExcelWorkSheet.Cells[1, 10] = "Litry AdBlue / 100 km";
                ExcelWorkSheet.Cells[1, 11] = "EUR Diesel / 100 km";
                ExcelWorkSheet.Cells[1, 12] = "EUR AdBlue / 100 km";
                ExcelWorkSheet.Cells[1, 13] = "Opłaty autostradowe / 100 km";
                ExcelWorkSheet.Cells[1, 14] = "Inne opłaty / 100 km";
                ExcelWorkSheet.Cells[1, 15] = "Wszystkie koszty / 100 km";

                foreach (Truck element in TruckData)
                {
                    String R = element.GetRegistration();
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = R;
                    CurrentColumn++;

                    double K = Math.Round((double)element.Kilometers, 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    K = Math.Round((double)element.AdblueCost, 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    K = Math.Round((double)element.AdblueL, 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    K = Math.Round((double)element.DieselCost, 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    K = Math.Round((double)element.DieselL, 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    K = Math.Round((double)element.RoadTax, 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    K = Math.Round((double)element.OtherCost, 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    //  PER 100

                    // LDiesel / 100 //TODO:
                    K = Math.Round((double)(element.DieselL / (float)element.Kilometers * 100.0f), 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    // LAdBlue / 100
                    K = Math.Round((double)(element.AdblueL / (float)element.Kilometers * 100.0f), 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    // Eur Diesel / 100
                    K = Math.Round((double)(element.DieselCost / (float)element.Kilometers * 100.0f), 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    // Eur AdBlue / 100
                    K = Math.Round((double)(element.AdblueCost / (float)element.Kilometers * 100.0f), 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    // RoadTax / 100
                    K = Math.Round((double)(element.RoadTax / (float)element.Kilometers * 100.0f), 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    // Other / 100
                    K = Math.Round((double)(element.OtherCost / (float)element.Kilometers * 100.0f), 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;
                    CurrentColumn++;

                    // All / 100
                    K = Math.Round((double)((

                        element.DieselL + element.DieselCost 
                        + element.AdblueL + element.AdblueCost
                        + element.OtherCost + element.RoadTax)

                        / (float)element.Kilometers * 100.0f), 2);
                    ExcelWorkSheet.Cells[CurrentRow, CurrentColumn] = K;

                    CurrentColumn = 1;
                    CurrentRow++;

                }

                //cosmetics
                ExcelWorkBook.Worksheets[1].Name = "Raport";//Renaming the Sheet1 to MySheet
                Microsoft.Office.Interop.Excel.Range aRange = ExcelWorkSheet.get_Range("A1", "O1");
                aRange.Columns.AutoFit();


                ExcelWorkBook.SaveAs(Path);
                ExcelWorkBook.Close(false);
                ExcelApp.Quit();
                //Marshal.ReleaseComObject(ExcelWorkSheet);
                //Marshal.ReleaseComObject(ExcelWorkBook);
                //Marshal.ReleaseComObject(ExcelApp);
            }
            //catch (Exception exHandle)
            {
                //Console.WriteLine("Exception: " + exHandle.Message);
                //Console.ReadLine();
            }
        }


        public void closeAll()
        {
            foreach( Microsoft.Office.Interop.Excel.Workbook singleWorkbook in MyExcel.Workbooks)
            {
                singleWorkbook.Close(false);
            }

            MyExcel.Quit();
        }
    }

}
