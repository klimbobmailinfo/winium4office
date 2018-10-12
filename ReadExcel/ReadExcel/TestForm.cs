using ClosedXML.Excel;
using OpenQA.Selenium;
using System;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;

namespace ReadExcel
{
    public partial class TestForm : Form
    {
        public TestForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            if (1 == 0)
            {

                OpenQA.Selenium.Winium.DesktopOptions desktopOptionsW = new OpenQA.Selenium.Winium.DesktopOptions();

                desktopOptionsW.ApplicationPath = @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE";

                OpenQA.Selenium.Winium.WiniumDriver driverW = new OpenQA.Selenium.Winium.WiniumDriver(new Uri("http://localhost:9999"), desktopOptionsW);

                System.Threading.Thread.Sleep(1000);

                var window = driverW.FindElement(By.Id("BackstageView"));

                window.SendKeys("Open Other Documents");
                //window.SendKeys(OpenQA.Selenium.Keys.Alt + "o");

                System.Threading.Thread.Sleep(1000);
                
                window = driverW.FindElement(By.Name("Open"));
                System.Threading.Thread.Sleep(1000);
                window.Click();
                System.Threading.Thread.Sleep(1000);

                window = driverW.FindElement(By.Name("Browse"));
                System.Threading.Thread.Sleep(1000);
                window.Click();
                System.Threading.Thread.Sleep(1000);

                window = driverW.FindElement(By.Name("File name:"));
                System.Threading.Thread.Sleep(1000);
                window.SendKeys(@"C:\netload\Artist Draft.docx");
                System.Threading.Thread.Sleep(1000);
                SendKeys.SendWait("{Enter}");

                //window = driverW.FindElement(By.Name("Open"));

                //window.SendKeys(OpenQA.Selenium.Keys.Alt + "o");
                /*
                window = driverW.FindElement(By.Id("1"));
                System.Threading.Thread.Sleep(1000);
                window.SendKeys(OpenQA.Selenium.Keys.Enter);
                System.Threading.Thread.Sleep(1000);
                */
                //                window.SendKeys("o");

                // driverW.
                foreach (var proc in Process.GetProcessesByName("Winium.Desktop.Driver"))
                {
                   // proc.Kill();
                }

                Environment.Exit(0);
            }

/*
            OpenQA.Selenium.Winium.DesktopOptions desktopOptions = new OpenQA.Selenium.Winium.DesktopOptions();

                desktopOptions.ApplicationPath = @"C:\Program Files (x86)\Microsoft Office\root\Office16\WinWord.EXE";

                OpenQA.Selenium.Winium.WiniumDriver driver = new OpenQA.Selenium.Winium.WiniumDriver(new Uri("http://localhost:9999"), desktopOptions);

                System.Threading.Thread.Sleep(4000);

            */


            string fileName;
            fileName = @"C:\Users\do8\Documents\Visual Studio 2015\Projects\ReadExcel\Mic Outlook Test.xlsx";
                //fileName = @"C:\Users\do8\Documents\Visual Studio 2015\Projects\ReadExcel\Mic Word Test.xlsx";
            /*
            using (XLWorkbook excelWorkbook = new XLWorkbook(fileName))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                foreach (var dataRow in nonEmptyDataRows)
                {
                    for (int i = 1; i < 3; i++)
                        Console.WriteLine(dataRow.Cell(i).Value);
                }
            }
            */
            var datatable = new DataTable();
            var workbook = new XLWorkbook(fileName);
            var xlWorksheet = workbook.Worksheet(1);
            var range = xlWorksheet.Range(xlWorksheet.FirstCellUsed(), xlWorksheet.LastCellUsed());

            var col = range.ColumnCount();
            var row = range.RowCount();

            //if a datatable already exists, clear the existing table 
            datatable.Clear();
            for (var i = 1; i <= col; i++)
            {
                var column = xlWorksheet.Cell(1, i);
                datatable.Columns.Add(column.Value.ToString());
            }

            var firstHeadRow = 0;
            foreach (var item in range.Rows())
            {
                if (firstHeadRow != 0)
                {
                    var array = new object[col];
                    for (var y = 1; y <= col; y++)
                    {
                        array[y - 1] = item.Cell(y).Value;
                    }

                    datatable.Rows.Add(array);
                }
                firstHeadRow++;
            }
            dataGridView1.DataSource = datatable;
            dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dataGridView1.Refresh();

            OpenQA.Selenium.Winium.DesktopOptions desktopOptionsE = new OpenQA.Selenium.Winium.DesktopOptions();
            OpenQA.Selenium.Winium.WiniumDriver driverE = null;
            IWebElement CurrentWindow = null;
            IWebElement Element = null;


            int maxRow = datatable.Rows.Count;
            int maxCol = datatable.Columns.Count;
            int cr = 1;
             // goto Auto;
            foreach (DataRow Row in datatable.Rows)
            {
                Console.Write("Row # " + cr++);
                for (int currentCol = 1; currentCol <= maxCol; currentCol++)
                {
                 //   Console.Write(currentCol + ": " + Row[currentCol - 1] + ",");
                }

                String Parameter = Row[2].ToString(); // Third Column
                String ToDo = Row[0].ToString().ToUpper(); // First Column

                Parameter = Parameter.Replace("[Date]", DateTime.Now.ToShortDateString());
                Parameter = Parameter.Replace("[Time]", DateTime.Now.ToShortTimeString());

                System.Threading.Thread.Sleep(500);

                switch (ToDo) // Iterate via first column
                {
                    case "DRIVER":
                        Console.WriteLine("Driver");
                        System.Diagnostics.Process.Start(Parameter); // if launch driver, the exe name is in 3rd column);
                        break;

                    case "STARTEXE":
                        Console.WriteLine("Run Exe: " + Parameter);
                        desktopOptionsE.ApplicationPath = Parameter; // For EXE, the exe name is in 3rd column
                        driverE = new OpenQA.Selenium.Winium.WiniumDriver(new Uri("http://localhost:9999"), desktopOptionsE);
                        System.Threading.Thread.Sleep(2000);
                        break;

                    case "FINDNAME":
                        Console.WriteLine("FindByName: " + Parameter);
                        Element = driverE.FindElement(By.Name(Parameter));
                        break;

                    case "FINDIDINOBJECT":
                        Console.WriteLine("FindIDInObject");
                        String[] IDs = Parameter.Split(';');
                        Element = (IWebElement)driverE.FindElement(By.Id(IDs[0])).FindElement(By.Id(IDs[1]));
                        CurrentWindow = Element;
                        break;

                    case "FINDID":
                        Console.WriteLine("FindByID: " + Parameter); 
                        Element = driverE.FindElement(By.Id(Parameter));
                        break;

                    case "PRESS":
                        Console.WriteLine("Press");
                        SendKeys.SendWait(Parameter);
                        break;

                    case "CLICK":
                        Console.WriteLine("Click");
                        Element.Click();
                        break;

                    case "TYPE":
                        Console.WriteLine("Send keys: " + Parameter);
                        Element.SendKeys(Parameter);
                        break;

                    case "END":
                        Console.WriteLine("End");
                        break;

                    default:
                        Console.WriteLine(ToDo + " --- ");
                        break;
                }


            }

            Auto:
            Environment.Exit(0);
            /*
            for(int currentRow = 1; currentRow<maxRow; currentRow++)
            {
                switch(datatable.Rows.IndexOf(currentRow)):
                datatable.R

            }
            */
            if (1==0)
            {


                OpenQA.Selenium.Winium.DesktopOptions desktopOptionsManual = new OpenQA.Selenium.Winium.DesktopOptions();

                desktopOptionsManual.ApplicationPath = @"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE";

                OpenQA.Selenium.Winium.WiniumDriver driverMa = new OpenQA.Selenium.Winium.WiniumDriver(new Uri("http://localhost:9999"), desktopOptionsManual);

                System.Threading.Thread.Sleep(4000);

                var mybutton = driverMa.FindElementByName("New Email");
                //var mybutton = driver.FindElementByXPath("//*[@Name='New Email']");
                //Name
                //mybutton.SendKeys("oh Dam!");
                System.Threading.Thread.Sleep(2000);
                mybutton.Click();

                System.Threading.Thread.Sleep(2000);

                var window = driverMa.FindElement(By.Id("258"));

                var subject = driverMa.FindElement(By.Id("258")).FindElement(By.Id("4101"));


                //var subject = window.FindElement(By.Id("4101"));

                //subject = driver.FindElementByXPath("//*[@Name='New Email']");

                var mytype = subject.GetAttribute("ClassName");

                if (mytype == "RichEdit20WPT")
                {
                    subject.SendKeys("Hi from Winium on " + DateTime.Now);
                }

                System.Threading.Thread.Sleep(2000);

                mybutton = window.FindElement(By.Id("4099")); // To:
                mybutton.SendKeys("thdam@wsgr.com");

                System.Threading.Thread.Sleep(2000);
                mybutton = window.FindElement(By.Id("Body")); // Body:
                mybutton.SendKeys("Sending my Regards to Paul McCartney!");

                System.Threading.Thread.Sleep(2000);
                mybutton = window.FindElement(By.Id("4256")); // Send
               // mybutton.Click();


            }




        }

        IWebElement RememberMe;

        private IWebElement PopDriver()
        {
            return RememberMe;
        }

        private void PushDriver(IWebElement currentWindow)
        {
            RememberMe = currentWindow;
        }
    }
}
