using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    enum ExcelType
    {
        TRUCK_DATA,
        EXPORT_GRID_DATA,
        JUST_NUMBERS,
        F_AND_NUMBERS,
        SN_AND_NUMBERS,
        ERROR
    }

    class ExcelRecogniser
    {
        public static ExcelType recognizeExcel(string pathToExcel)
        {
            
            string extension = System.IO.Path.GetExtension(pathToExcel).ToLower();
            if (pathToExcel.Contains( "SN760756.xlsx"))
                return ExcelType.SN_AND_NUMBERS;

            else if (pathToExcel.Contains("ExportGridData"))
                return ExcelType.EXPORT_GRID_DATA;

            else if (pathToExcel.Contains("Monatsinfos"))
                return ExcelType.TRUCK_DATA;


            if (!extension.Contains("xls")) {
                System.Windows.Forms.MessageBox.Show(
                    "Cannot recognise " + pathToExcel + ". Are you sure the file is valid?"
                    );
                return ExcelType.ERROR;
            }


            Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
            Microsoft.Office.Interop.Excel.Range MyCells;
            Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook MyWorkbook;

            MyWorkbook = MyExcel.Workbooks.Open(pathToExcel);
            MyWorksheet = MyExcel.Worksheets.Item[1];
            MyCells = MyWorksheet.Cells;


            string firstCell = (System.Convert.ToString(MyCells.Item[1, 1].Value)) + "";
            MyWorkbook.Close(false);
            MyExcel.Quit();


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

                default:
                    return ExcelType.ERROR;

            }


            return ExcelType.ERROR;

        }
    }
}
