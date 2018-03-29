using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace EZ_Finance
{





    class WriteToExcel
    {
        public int currentSheet { get; set; }
        //public string templatePath { get; set; }
        public static BindingList<DCU_data> DCUDataList = new BindingList<DCU_data>();
        public static BindingList<NET_data> NETDataList = new BindingList<NET_data>();
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static int lastRow = 0;


    public void InitializeExcel(string templatepath)

        {
            
            MyApp = new Excel.Application();
            MyApp.Visible = true;
            MyBook = MyApp.Workbooks.Open(templatepath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[currentSheet];
            //lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            lastRow = 3;
        }

   

    public void ChangeSheet(int currentSheet)
        {
            MySheet = (Excel.Worksheet)MyBook.Sheets[currentSheet];
            lastRow = 3;
        }

    public static BindingList<DCU_data> ReadMyExcelDCU()
        {
            DCUDataList.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "E" + index.ToString()).Cells.Value;
                DCUDataList.Add(new DCU_data
                {
                    Date = MyValues.GetValue(1, 1).ToString(),
                    Description = MyValues.GetValue(1, 2).ToString(),
                    Deposit = MyValues.GetValue(1, 3).ToString(),
                    Withdrawl = MyValues.GetValue(1, 4).ToString(),
                    Balance = MyValues.GetValue(1, 5).ToString()
                });
            }
            return DCUDataList;


        }


        public static void writeAccountsToExcelDCU(DCU_data accounts)
        {    
                MySheet.Cells[3, 5] = accounts.Checking;
                MySheet.Cells[4, 5] = accounts.Savings;
                MySheet.Cells[6, 5] = "-" + accounts.Credit;
        }


        public static void writeTransactionsToExcelDCU (DCU_data transactions)
        {
            try
            {
                lastRow += 1;
                MySheet.Cells[lastRow, 1] = transactions.Date;
                MySheet.Cells[lastRow, 2] = transactions.Description;
                MySheet.Cells[lastRow, 3] = transactions.Deposit;

            }
            catch (Exception)
            { }

            
            //MyBook.Close();
            //MyApp.Quit();

        }


        public static BindingList<NET_data> ReadMyExcelNET()
        {
            NETDataList.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "E" + index.ToString()).Cells.Value;
                NETDataList.Add(new NET_data
                {
                    Date = MyValues.GetValue(1, 1).ToString(),
                    Description = MyValues.GetValue(1, 2).ToString(),
                    Amount = MyValues.GetValue(1, 3).ToString(),
                    Balance = MyValues.GetValue(1, 4).ToString(),
                    
                });
            }
            return NETDataList;


        }

        public static void writeAccountsToExcelNET(NET_data accounts)
        {
            MySheet.Cells[3, 2] = accounts.Checking;
            MySheet.Cells[4, 2] = accounts.Savings;
        }

        public static void writeTransactionsToExcelNET(NET_data transactions)
        {
            try
            {
                lastRow += 1;
                MySheet.Cells[lastRow, 1] = transactions.Date;
                MySheet.Cells[lastRow, 2] = transactions.Description;
                MySheet.Cells[lastRow, 3] = transactions.Amount;
                MySheet.Cells[lastRow, 4] = transactions.Balance;
                
                //NETDataList.Add(transactions);


            }
            catch (Exception)
            { }


            //MyBook.Close();
            //MyApp.Quit();

        }
    

        public static void saveExcelFile(FileName_data fileName)
        {
            MyApp.DisplayAlerts = false;
            MyBook.SaveAs(fileName.excelPath + fileName.currentUser + "-" + fileName.currentDate + ".xlsx");
        }

        public void Cleanup()
        {

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(MySheet);

            //close and release
            MyBook.Close();
            Marshal.ReleaseComObject(MyBook);

            //quit and release
            MyApp.Quit();
            Marshal.ReleaseComObject(MyApp);



        }

    }
}
