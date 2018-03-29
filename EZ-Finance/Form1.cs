using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace EZ_Finance
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }


        //Decalred varaibles that interact with all of main
        string currentDate, currentUser = "RZELLER", sendAddress = "rogerjohnmorellizeller@gmail.com", senderEmailAddress, senderEmailPassword;
        string excelPath = @"F:\Documents\Expenses\", userDataPath = @"F:\Documents\Expenses\UserData.xlsx", templateName = "Finances Template.xlsx";
        string NET_User, NET_Pass, DCU_User, DCU_Pass;
        int currentSheet;
        
        //Send data from labels to bank data for main Accounts 
        WriteToExcel startExcel = new WriteToExcel();
        ReadUserData readData = new ReadUserData();
        
        //Executes commands on form load 
        private void Form1_Load(object sender, EventArgs e)
        {
           
            
        }

        //Starts ripping data from bank pages and saves to excel file
        private void btnStart_Click(object sender, EventArgs e)
        {

            
            obtainCurrentDate();
            obtainUserData();
            initializeExcel();
            //NETlogin();
            DCUlogin();
            //GoogleLogin();
            //firefoxtest();
            //endProccesses();
            saveExcel();
            closeExcel();
            sendEmail();
            startExcel.Cleanup();
            this.Close();
            

        }

        //Saves excel data when clicked. *Not in use* 
        private void btnSave_Click(object sender, EventArgs e)
        {
            obtainUserData();
            
        }

        //Obtains current date. Needed for nameing and saving files correctly
        private void obtainCurrentDate()
        {
            DateTime localDate = DateTime.Now;
            currentDate = DateTime.Now.ToString("M-d-yyyy");
        }

        //Obtains Username and Password data for both banks
        private void obtainUserData()
        {
            readData.currentSheet = 1;
            readData.InitializeExcel(userDataPath);
            //readData.obtainExcelDataTest();

            //Obtain NET Data 
            NET_data NET_UserInfo = new NET_data();
            readData.obtainDataNET(NET_UserInfo);
            NET_User = NET_UserInfo.User;
            NET_Pass = NET_UserInfo.Pass;


            //Obtain DCU Data
            DCU_data DCU_UserInfo = new DCU_data();
            readData.obtainDataDCU(DCU_UserInfo);
            DCU_User = DCU_UserInfo.User;
            DCU_Pass = DCU_UserInfo.Pass;

            //Obtain Email Sender Data 
            MailSender_data MailSender_UserInfo = new MailSender_data();
            readData.obtainSendEmailData(MailSender_UserInfo);
            senderEmailAddress = MailSender_UserInfo.senderMailAddress;
            senderEmailPassword = MailSender_UserInfo.senderMailPassword;

            //Cleanup excel app 
            readData.Cleanup();

        }

        //Open excel, set current data sheet, and set template path 
        public void initializeExcel() 
        {
            //Open excel and set accounts excel sheet 
            startExcel.currentSheet = 2;
            startExcel.InitializeExcel(excelPath + templateName);
        }

        //Save to excel file
        public void saveExcel()
        {
            //set current date to be written as excel file name 
            FileName_data fileName = new FileName_data()
            {
                currentDate = currentDate,
                currentUser = currentUser,
                excelPath = excelPath
            };

            WriteToExcel.saveExcelFile(fileName);
            System.Threading.Thread.Sleep(10000);
        }

        //close excel when done writing to file 
        public static void closeExcel()
        {
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
                System.Threading.Thread.Sleep(10000);
            }
        }

       
        //NET Federal Credit Union Login 
        private void NETlogin()
        {
            //NET Federal Bank Login 

            //Run selenium
            FirefoxDriver fd = new FirefoxDriver();
            fd.Url = @"https://netfedcu.online-cu.com/ISuite5/Features/Auth/MFA/Default.aspx";
            fd.Navigate();
            IWebElement r = fd.FindElementById("ctl01_Main1_UserIDTextbox");
            r.SendKeys(NET_User);
            r = fd.FindElementById("ctl01_Main1_LoginBtn");
            r.Click();
            new WebDriverWait(fd, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("ctl01_Main1_PasswordTextbox"))));
            r = fd.FindElementById("ctl01_Main1_PasswordTextbox");
            r.SendKeys(NET_Pass);
            r = fd.FindElementById("ctl01_Main1_SignInBtn");
            r.Click();

            //Get Data for Checking 
            new WebDriverWait(fd, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("ctl01_Main1_MemberTab_objc0_ctl00_AccountSummaryLiteGrid_ctl03_AvailBalLabel"))));
            r = fd.FindElementById("ctl01_Main1_MemberTab_objc0_ctl00_AccountSummaryLiteGrid_ctl03_AvailBalLabel");
            lblNetCheckingData.Text = r.GetAttribute("innerHTML");

            //Get Data for Savings 
            new WebDriverWait(fd, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("ctl01_Main1_MemberTab_objc0_ctl00_AccountSummaryLiteGrid_ctl02_AvailBalLnkBtn"))));
            r = fd.FindElementById("ctl01_Main1_MemberTab_objc0_ctl00_AccountSummaryLiteGrid_ctl02_AvailBalLnkBtn");
            lblNetSavingsData.Text = r.GetAttribute("innerHTML");

            //Store accounts data in NET_data
            NET_data accounts = new NET_data()
            {
                Checking = lblNetCheckingData.Text,
                Savings = lblNetSavingsData.Text,
            };

            //Write Checking and Savings to excel file 
            WriteToExcel.writeAccountsToExcelNET(accounts);

            //Change Sheet to NET checking 
            startExcel.currentSheet = 5;

            //navigate to checking table
            //r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_MemberTab_objc0_ctl00_AccountSummaryLiteGrid_ctl03_ACSLabelLnkBtn']"));
            r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_MemberTab_objc0_ctl00_AccountSummaryLiteGrid_ctl03_RecentActivityLnkBtn']"));
            r.Click();
            System.Threading.Thread.Sleep(10000);


            //Select show last 28 days in history combo box
            IWebElement iWebelement = fd.FindElement(By.Id("ctl01_Main1_dropdownDateRange")); //Getting the element of the Text Box
            IWebElement iWebelementList = fd.FindElement(By.Id("ctl01_Main1_dropdownDateRange")); //Getting the elements of List/Drop down Box
            SelectElement selected = new SelectElement(iWebelementList); //Parsing the list

            iWebelement.SendKeys(OpenQA.Selenium.Keys.ArrowDown); // Clicking the drop down image

            selected.SelectByText("Last 28 Days"); // Use select element class to select the value

            //click show history button 
            r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_UISButtonShowHistory']")); 
            r.Click();
            System.Threading.Thread.Sleep(10000);

           
            //Check for number of rows to obtain data from 
            for (int x = 2; x < 1000;)
            {

                // Store detected elements in a data list
                List<IWebElement> elementList = new List<IWebElement>();
                elementList.AddRange(fd.FindElements(By.XPath(".//*[@id='ctl01_Main1_SufixListRepeater_ctl03_HistoryForRepeatSuffix1_SFXHistoryRollup_SFXHistoryGrid']/tbody/tr[" + x + "]/td[1]")));

                if (elementList.Count > 0)
                {
                    x++;
                    lbltest.Text = x.ToString();
                }
                else
                {
                    break;
                }

            }
            //Convert element count to an integer. Decrease by 2 for accuracy when looping. 
            int checkEleCount = Convert.ToInt32(lbltest.Text);
            checkEleCount = checkEleCount - 2;

            //change to checking excel sheet 
            currentSheet = 5;
            startExcel.ChangeSheet(currentSheet);


            //Add data for each transaction
            for (int x = 1; x <= checkEleCount;)


            {
                Label[] activeLabel = { lblNetDateRes,
                    lblNetDesRes,
                    lblNetAmtRes,
                    lblWithRes
            };

                // Default strings. Changes after assigned new value. 
                string activeNetDate = "No good Date";
                string activeNetDes = "No good Des";
                string activeNetAmount = "No good Amt"; 
                string activeNetBalance = "No good Bal";
                string[] currentElement = { activeNetDate, activeNetDes, activeNetAmount, activeNetBalance};

                //assign values to labels for each part of the transaction in the selected row 
                {

                    r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_SufixListRepeater_ctl03_HistoryForRepeatSuffix1_SFXHistoryRollup_SFXHistoryGrid']/tbody/tr[" + (x + 1) + "]/td[1]"));
                    activeLabel[0].Text = r.GetAttribute("innerHTML");
                    currentElement[0] = activeLabel[0].Text;


                    if (x < 9)
                    {
                        r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_SufixListRepeater_ctl03_HistoryForRepeatSuffix1_SFXHistoryRollup_SFXHistoryGrid_ctl0" + (x + 1) + "_UneditedDescription']"));
                    }
                    else
                    { 
                    r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_SufixListRepeater_ctl03_HistoryForRepeatSuffix1_SFXHistoryRollup_SFXHistoryGrid_ctl" + (x + 1) + "_UneditedDescription']"));
                    
               }

                    
                    activeLabel[1].Text = r.GetAttribute("innerHTML");
                    currentElement[1] = activeLabel[1].Text;

                    r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_SufixListRepeater_ctl03_HistoryForRepeatSuffix1_SFXHistoryRollup_SFXHistoryGrid']/ tbody/tr[" + (x + 1) + "]/td[3]"));
                    activeLabel[2].Text = r.GetAttribute("innerHTML");
                    currentElement[2] = activeLabel[2].Text;

                    r = fd.FindElement(By.XPath(".//*[@id='ctl01_Main1_SufixListRepeater_ctl03_HistoryForRepeatSuffix1_SFXHistoryRollup_SFXHistoryGrid']/ tbody/tr[" + (x + 1) + "]/td[4]"));
                    activeLabel[3].Text = r.GetAttribute("innerHTML");
                    currentElement[3] = activeLabel[3].Text;


                }

                NET_data transactions = new NET_data
                {
                    Date = currentElement[0],
                    Description = currentElement[1],
                    Amount = currentElement[2],
                    Balance = currentElement[3],
                    
                };

                //Write transactions to excel 
                WriteToExcel.writeTransactionsToExcelNET(transactions);
                x++;

            }

            fd.Navigate().Back();




            //Log out of NET Bank 
            new WebDriverWait(fd, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("ctl01_LogoutImage"))));
            r = fd.FindElementById("ctl01_LogoutImage");
            r.Click();

            //Send data from labels to NET_data for main Accounts 
            //MyExcel netExcel = new MyExcel();

            ////Open excel and set accounts excel sheet 
            //netExcel.currentSheet = 2;
            //netExcel.InitializeExcel(excelPath + templateName);







            

            

            System.Threading.Thread.Sleep(10000);
            new WebDriverWait(fd, TimeSpan.FromSeconds(120));
            fd.Close();
            System.Threading.Thread.Sleep(10000);
            fd.Quit();
        }

        //DCU Login 
        private void DCUlogin()
        {



            //Open Chrome window
            ChromeOptions options = new ChromeOptions();

            options.AddArgument("user-data-dir=C:/Users/RJ's Desktop/AppData/Local/Google/Chrome/User Data/Profile 1");
            var cd = new ChromeDriver(options);

            //naviagte to DCU.org and login 
            cd.Url = @"https://www.dcu.org/";
            cd.Navigate();
            IWebElement r = cd.FindElementById("userid");
            r.SendKeys(DCU_User);
            r = cd.FindElementById("password");
            r.SendKeys(DCU_Pass);
            r = cd.FindElementById("submitBtn");
            r.Click();

            //MessageBox.Show("derp wait");
            cd.SwitchTo().Frame("appContainer");
            //Get Data for Checking
            new WebDriverWait(cd, TimeSpan.FromSeconds(20));
            new WebDriverWait(cd, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.XPath(".//*[@id='ext-comp-1014']/div[1]/div[2]/div[1]/div[2]/span"))));
            r = cd.FindElement(By.XPath(".//*[@id='ext-comp-1014']/div[1]/div[2]/div[1]/div[2]/span"));
            lblDCUCheckingData.Text = r.GetAttribute("innerHTML");





            //Get Data for Savings 
            r = cd.FindElement(By.XPath(".//*[@id='ext-comp-1014']/div[2]/div[2]/div[1]/div[2]/span"));
            lblDCUSavingsData.Text = r.GetAttribute("innerHTML");

            //Get Data for VISA Credit Card 
            r = cd.FindElement(By.XPath(".//*[@id='ext-comp-1023']/div/div[2]/div[1]/div[2]/span"));
            lblDCUCreditData.Text = r.GetAttribute("innerHTML");

            //Change to account sheet for dcu 
            WriteToExcel dcuExcel = new WriteToExcel();

            currentSheet = 2;
            dcuExcel.ChangeSheet(currentSheet);

            //Store accounts data in DCU_data
            DCU_data accounts = new DCU_data()
            {
                Checking = lblDCUCheckingData.Text,
                Savings = lblDCUSavingsData.Text,
                Credit = lblDCUCreditData.Text

            };

            WriteToExcel.writeAccountsToExcelDCU(accounts);


            //navigate to checking table
            r = cd.FindElement(By.XPath(".//*[@id='ext-gen43']/span"));
            r.Click();


            //cd.SwitchTo().Frame("appContainer");
            new WebDriverWait(cd, TimeSpan.FromSeconds(60)).Until(ExpectedConditions.ElementExists((By.XPath(".//*[@id='lblAccountGrid']/span/h3"))));



            //obtain data 
            IWebElement test = cd.FindElement(By.XPath(".//*[@id='lblAccountGrid']/span/h3"));
            System.Threading.Thread.Sleep(10000);

            //Check for number of rows to obtain data from 
            for (int x = 1; x < 1000;)
            {

                // Store detected elements in a data list
                List<IWebElement> elementList = new List<IWebElement>();
                elementList.AddRange(cd.FindElements(By.XPath(".//*[@id='ext-gen34']/table/tbody/tr[" + x + "]/td[2]/div")));

                if (elementList.Count > 0)
                {
                    x++;
                    lbltest.Text = x.ToString();
                }
                else
                {
                    break;
                }

            }
            //Convert element count to an integer. Decrease by 1 for accuracy when looping. 
            int checkEleCount = Convert.ToInt32(lbltest.Text);
            checkEleCount--;

            //change to checking excel sheet 
            currentSheet = 3;
            dcuExcel.ChangeSheet(currentSheet);


            //Add data for each transaction
            for (int x = 1; x <= checkEleCount;)


            {



                Label[] activeLabel = { lblDateRes,
                    lblDesRes,
                    lblDeposRes,
                    lblWithRes,
                    lblBalRes, };

                // Default strings. Changes after assigned new value. 
                string activeDate = "No good";
                string activeDes = "No good";
                string activeDepos = "No good";
                string activeWith = "No good";
                string activeBal = "No good";
                string[] currentElement = { activeDate, activeDes, activeDepos, activeWith, activeBal };

                //loop through each part of a row for each transaction 
                for (int y = 2; y < 7;)

                {

                    try


                    {
                        r = cd.FindElement(By.XPath(".//*[@id='ext-gen34']/table/tbody/tr[" + x + "]/td[" + y + "]/div"));
                        activeLabel[y - 2].Text = r.GetAttribute("innerHTML");
                        currentElement[y - 2] = activeLabel[y - 2].Text;




                    }
                    catch (NoSuchElementException)
                    {
                        r = cd.FindElement(By.XPath(".//*[@id='ext-gen34']/table/tbody/tr[" + x + "]/td[" + y + "]"));
                        activeLabel[y - 2].Text = r.GetAttribute("innerHTML");
                        currentElement[y - 2] = activeLabel[y - 2].Text;


                    }


                    //go to next column in the row 
                    y++;



                }

                DCU_data transactions = new DCU_data
                {
                    Date = currentElement[0],
                    Description = currentElement[1],
                    Deposit = currentElement[2],
                    Withdrawl = currentElement[3],
                    Balance = currentElement[4],
                };

                WriteToExcel.writeTransactionsToExcelDCU(transactions);
                //MessageBox.Show("Details were successfully added to the excel file!", "Success..", MessageBoxButtons.OK, MessageBoxIcon.Information);
                x++;

            }

            cd.Navigate().Back();




            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////Credit/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //MessageBox.Show("derp wait");
            cd.SwitchTo().Frame("appContainer");



            new WebDriverWait(cd, TimeSpan.FromSeconds(20));
            new WebDriverWait(cd, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.XPath(".//*[@id='ext-comp-1014']/div[1]/div[2]/div[1]/div[2]/span"))));
            r = cd.FindElement(By.XPath(".//*[@id='ext-comp-1014']/div[1]/div[2]/div[1]/div[2]/span"));


            //navigate to credit table
            r = cd.FindElement(By.XPath(".//*[@id='ext-gen55']/span"));
            r.Click();

        
            //cd.SwitchTo().Frame("appContainer");
            new WebDriverWait(cd, TimeSpan.FromSeconds(60)).Until(ExpectedConditions.ElementExists((By.XPath(".//*[@id='lblAccountGrid']/span/h3"))));



            //obtain data 
            IWebElement testCredit = cd.FindElement(By.XPath(".//*[@id='lblAccountGrid']/span/h3"));
            System.Threading.Thread.Sleep(10000);

            //Check for number of rows to obtain data from. This will rarely ever exceed 1000 per month. 
            for (int x = 1; x < 1000;)
            {

                // Store detected elements in a data list
                List<IWebElement> elementList = new List<IWebElement>();
                elementList.AddRange(cd.FindElements(By.XPath(".//*[@id='ext-gen34']/table/tbody/tr[" + x + "]/td[2]/div")));

                if (elementList.Count > 0)
                {
                    x++;
                    lbltest.Text = x.ToString();
                }
                else
                {
                    break;
                }

            }

            //Convert element count to an integer. Decrease by 1 for accuracy when looping
            int checkEleCountCredit = Convert.ToInt32(lbltest.Text);
            checkEleCountCredit--;


            //Change to credit dcu excel sheet 
            WriteToExcel dcuExcelCredit = new WriteToExcel();

            currentSheet = 4;
            dcuExcelCredit.ChangeSheet(currentSheet);

            //Add data for each transaction
            for (int x = 1; x <= checkEleCountCredit;)


            {



                Label[] activeLabel = { lblDateRes,
                    lblDesRes,
                    lblDeposRes,
                    lblWithRes,
                    lblBalRes, };

                // Default strings. Changes after assigned new value
                string activeDate = "No good";
                string activeDes = "No good";
                string activeDepos = "No good";
                string activeWith = "No good";
                string activeBal = "No good";
                string[] currentElement = { activeDate, activeDes, activeDepos, activeWith, activeBal };

                //loop through each part of a row for each transaction 
                for (int y = 2; y < 7;)

                {

                    try


                    {
                        r = cd.FindElement(By.XPath(".//*[@id='ext-gen34']/table/tbody/tr[" + x + "]/td[" + y + "]/div"));
                        activeLabel[y - 2].Text = r.GetAttribute("innerHTML");
                        currentElement[y - 2] = activeLabel[y - 2].Text;




                    }
                    catch (NoSuchElementException)
                    {
                        r = cd.FindElement(By.XPath(".//*[@id='ext-gen34']/table/tbody/tr[" + x + "]/td[" + y + "]"));
                        activeLabel[y - 2].Text = r.GetAttribute("innerHTML");
                        currentElement[y - 2] = activeLabel[y - 2].Text;


                    }



                    y++;



                }

                DCU_data transactions = new DCU_data
                {
                    Date = currentElement[0],
                    Description = currentElement[1],
                    Deposit = currentElement[2],
                    Withdrawl = currentElement[3],
                    Balance = currentElement[4],
                };

                WriteToExcel.writeTransactionsToExcelDCU(transactions);

                x++;






                //for (int data = 0; data < 5;)
                //{
                //    if (activeData[data] == "&nbsp;")
                //    {
                //        activeData[data] = "";
                //    }

                //    data++;




            }

            ////set current date to be written as excel file name 
            //FileName fileName = new FileName()
            //{
            //    currentDate = currentDate,
            //    currentUser = currentUser,
            //    excelPath = excelPath
            //};

            //MyExcel.saveExcelFile(fileName);


            cd.Navigate().Back();
            r = cd.FindElement(By.XPath(".//*[@id='logout']"));
            r.Click();



            System.Threading.Thread.Sleep(10000);
            new WebDriverWait(cd, TimeSpan.FromSeconds(120));
            new WebDriverWait(cd, TimeSpan.FromSeconds(60)).Until(ExpectedConditions.ElementExists((By.XPath(".//*[@id='userid']"))));
            cd.Close();
            System.Threading.Thread.Sleep(10000);
            cd.Quit();
        }


        //Test for google login 
        private void GoogleLogin()
        {
            //NET Federal Bank Login 

            //Run selenium
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("user-data-dir=C:/Users/RJ's Desktop/AppData/Local/Google/Chrome/User Data/Profile 1");
            var cd = new ChromeDriver(options);
            cd.Url = @"https://www.dcu.org/";
            cd.Navigate();
            IWebElement r = cd.FindElement(By.XPath("html/body/div[1]/div[2]/aside/div/div[2]/p[1]/a"));
            //r.SendKeys("derp");
            lblNetSavingsData.Text = r.GetAttribute("innerHTML");

        }

        //alternative firefox dcu login
        private void firefoxtest()
        {

            var cd = new FirefoxDriver();
            cd.Url = @"https://www.dcu.org/";
            cd.Navigate();
            IWebElement r = cd.FindElementById("userid");
            r.SendKeys("5821077");
            r = cd.FindElementById("password");
            r.SendKeys("Zsaber6660##");
            //r = cd.FindElementById("submitBtn");
            r.FindElement(By.XPath(".//*[@id='submitBtn']"));

            r.Click();
        }

        //End Gecko.exe
        private void endProccesses()
        {





            foreach (var process in Process.GetProcessesByName("geckodriver"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("IEDriverServer"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("chromedriver"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("firefox"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("iexplore"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("chrome"))
            {
                process.Kill();
            }

            //DCU 5821077
        }

        //Maxmize window 
        private void OpenURL(IWebDriver AppDriver)
        {
            try
            {
                AppDriver.Manage().Window.Maximize();
                AppDriver.SwitchTo().ActiveElement();
            }
            catch (Exception e)
            {
                Console.WriteLine("ERR: {0}; {1}", e.TargetSite, e.Message);
                throw;
            }
        }

        //Used to minimize cmdline 
        private void minimizeProcess()
        {




        }

        //removes empty strings from showing in excel sheet
        public static void filterEmptyString(string[] array)

        {
            for (int x = 0; x < array.Length;)
            {
                if (array[x] == "&nbsp;")
                {
                    array[x] = "";
                }

                x++;
            }
        }

       //Send email to default email address 
        private void sendEmail()
        {

            // assign values to mail method 
            SendMail mail = new SendMail()
            {
                fromAddressA = "abc@mydomain.com",
                fromAddressB = "EZ Finance",
                toAddress = sendAddress,
                subjectText = "Daily financial report" + currentDate,
                bodyText = "Please see attachment for daily financial report.",
                attachmentPath = excelPath + currentUser + "-" + currentDate + ".xlsx",
                usernameMail = senderEmailAddress,
                passwordMail = senderEmailPassword,
            };

            mail.sendGmail();





        }

    }


}
