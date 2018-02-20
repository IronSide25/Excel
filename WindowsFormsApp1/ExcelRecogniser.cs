using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    enum ExcelType
    {
        TRUCK_DATA, //check
        EXPORT_GRID_DATA, //check
        JUST_NUMBERS, //check
        F_AND_NUMBERS, //check
        SN_AND_NUMBERS, //check
        EXTRA_INVOICE, //extra check
        ERROR //shh, pls don't exist
    }

    class ExcelRecogniser
    {
        public static ExcelType recognizeExcel(string pathToExcel)
        {

            string extension = System.IO.Path.GetExtension(pathToExcel).ToLower();
            if (extension.Contains( "csv"))
                return ExcelType.SN_AND_NUMBERS;

            // should this exist?
            //if (System.IO.Path.GetFileName(pathToExcel).Contains("SN"))
            //    return ExcelType.SN_AND_NUMBERS;

            // nah

            
            // but this should exist
            if (System.IO.Path.GetFileName(pathToExcel).ToLower().Contains("extra"))
                return ExcelType.EXTRA_INVOICE;


            else if (pathToExcel.Contains("ExportGridData"))
                return ExcelType.EXPORT_GRID_DATA;

            else if (pathToExcel.Contains("Monatsinfos"))
                return ExcelType.TRUCK_DATA;



            string firstCell = "";
            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;
            Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook MyWorkbook;
            try { 
            MyWorkbook = MyExcel.Workbooks.Open(pathToExcel);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;
           

            firstCell = (System.Convert.ToString(MyCells.Item[1, 1].Value)) + "";
            MyWorkbook.Close(false);
            MyExcel.Quit();

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Nie udało się otworzyć " + pathToExcel);
                return ExcelType.ERROR;
            }
             


            switch (firstCell)
            {
                case "Numer Karty":
                    return ExcelType.EXPORT_GRID_DATA;
                        break;
                case "Umowa":
                    return ExcelType.F_AND_NUMBERS;
                    break;
                case "Index":
                    return ExcelType.JUST_NUMBERS;
                    break;
                case "":
                    return ExcelType.TRUCK_DATA;
                    break;
                case "Karte;Kennzeichen;Fahrer;km-Stand;Lieferdatum;Lieferzeit;Beleg-Nr;Erfassungsart;Land;SST;Name;Warenart;Menge;Preis;Betrag;Nachlass incl. USt;Fee incl. USt;Wert incl. USt;;Wert incl. USt;;USt;Betrag USt;Nachlass excl. USt;Fee excl. USt;Wert excl. USt;;Wert excl. USt;;Rechnung;Kostenstelle;":
                    return ExcelType.SN_AND_NUMBERS;
                    //but this should never happen
                    break;
                case "Rejestracja":
                    return ExcelType.EXTRA_INVOICE;
                    break;

                default:
                    return ExcelType.ERROR;

            }


            return ExcelType.ERROR;

        }
    }
}
