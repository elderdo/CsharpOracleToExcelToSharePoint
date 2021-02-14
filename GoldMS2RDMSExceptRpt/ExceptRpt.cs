using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace GoldMS2RDMSExceptRpt
{

    class ExceptRpt
    {
        private int headerRow = 2;

        private void setHeader(Excel.Worksheet xlWorkSheet, String range,
            String header, long color)
        {
            xlWorkSheet.Range[range].Interior.Color = color;
            xlWorkSheet.Range[range].Merge();
            xlWorkSheet.Range[range].Value = header;
            xlWorkSheet.Range[range].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

        }
        private void rowOneHdr(Excel.Worksheet xlWorkSheet)
        {
            xlWorkSheet.Range["A1:B1"].Merge();
            DateTime mountain = TimeZoneInfo.ConvertTimeBySystemTimeZoneId
           (DateTime.Now, "Mountain Standard Time");
            xlWorkSheet.Range["A1:B1"].Value = String.Format("{0:M/d/yyyy h:mm:ss tt} MST", mountain);

            long yellow = 13434879;
            setHeader(xlWorkSheet, "C1:K1", "Mesa GOLD", yellow);

            long blue = 16777164;
            setHeader(xlWorkSheet, "L1:S1", "MS²--Mesa", blue);

            long green = 13434828;
            setHeader(xlWorkSheet, "T1:Y1", "RDMS", green);

            long ugly_orange = 10079487;
            setHeader(xlWorkSheet, "Z1:AG1", "RDMS shipper", ugly_orange);
        }

        private void rowTwoHdr(Excel._Worksheet xlWorkSheet, DataSet ds)
        {
            long gray = 12632256;
            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
            {
                xlWorkSheet.Cells[headerRow, i + 1] = ds.Tables[0].Columns[i].ColumnName;
                xlWorkSheet.Cells[headerRow, i + 1].Interior.Color = gray;
            }

        }

        private void formatSheet(Excel.Worksheet xlWorkSheet)
        {
            String numFmt = "0";

            xlWorkSheet.Columns["A", Type.Missing].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            String dateFmt = "m/d/yyyy";
            xlWorkSheet.Columns["AD:AD", Type.Missing].NumberFormat = dateFmt;

            String[] numCols = {"AF:AF","AC:AC","AB:AB","AA:AA","V:Y","N:S","FJ"} ;
            foreach (String rng in numCols)
            {
                xlWorkSheet.Columns["AF:AF", Type.Missing].NumberFormat = numFmt;
            }

            String textFmt = "@";

            xlWorkSheet.Columns["AG", Type.Missing].NumberFormat = textFmt;

        }

        private void enterData(Excel.Worksheet xlWorkSheet,DataSet ds)
        {
            string data = null;
            for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
            {
                for (int j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[i + headerRow + 1, j + 1] = data;
                }
            }

        }

        public void createExceptRpt(DataSet ds, String fileName)
        {
            

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            rowOneHdr(xlWorkSheet);
            rowTwoHdr(xlWorkSheet,ds);
            formatSheet(xlWorkSheet);
            enterData(xlWorkSheet, ds);
            // make sure all the data fits
            xlWorkSheet.Columns["A:AG", Type.Missing].EntireColumn.AutoFit();


            xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            Logger.Info("Excel file created: " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + fileName,"ExceptRpt");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Logger.Error(ex, "ExceptRpt");
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
