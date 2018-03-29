using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace EZ_Finance
{
    class ReadUserData
    {

        //Used to test accurately obtaining data for user names and password for both accounts respectively. 
        string Net_userTest, Net_passTest,
        DCU_userTest, DCU_passTest, Sender_userTest, Sender_passTest; 
        Double DCU_userDouble;

        public int currentSheet { get; set; }
        private static Excel.Workbook Book = null;
        private static Excel.Application App = null;
        private static Excel.Worksheet Sheet = null;
        public void InitializeExcel(string templatepath)

        {

            App = new Excel.Application();
            App.Visible = true;
            Book = App.Workbooks.Open(templatepath);
            Sheet = (Excel.Worksheet)Book.Sheets[currentSheet];

        }

      
        public void obtainExcelDataTest()
        {
            
            Net_userTest = Sheet.Cells[2, 2].value();
            Net_passTest = Sheet.Cells[3, 2].value();
            // *Note* Data is collected as a double and then converted to a string before being assigned to User variable. 
            DCU_userDouble = Sheet.Cells[2, 3].value();
            DCU_userTest = DCU_userDouble.ToString();
            DCU_passTest = Sheet.Cells[3, 3].value();
            Sender_userTest = Sheet.Cells[2, 4].value();
            Sender_passTest = Sheet.Cells[3, 4].value();

        }

        public void obtainDataNET(NET_data UserInfo)
        {
            UserInfo.User = Sheet.Cells[2, 2].value();
            UserInfo.Pass = Sheet.Cells[3, 2].value(); 

        }

        public void obtainDataDCU(DCU_data UserInfo)
        {
            DCU_userDouble = Sheet.Cells[2, 3].value();
            // *Note* Data is collected as a double and then converted to a string before being assigned to User variable. 
            UserInfo.User = DCU_userDouble.ToString();
            UserInfo.Pass = Sheet.Cells[3, 3].value();
        }

        public void obtainSendEmailData(MailSender_data UserInfo)
        {
            UserInfo.senderMailAddress = Sheet.Cells[2, 4].value();
            UserInfo.senderMailPassword = Sheet.Cells[3, 4].value();
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
            Marshal.ReleaseComObject(Sheet);

            //close and release
            Book.Close();
            Marshal.ReleaseComObject(Book);

            //quit and release
            App.Quit();
            Marshal.ReleaseComObject(App);
            


        }
    }
}
        
    


