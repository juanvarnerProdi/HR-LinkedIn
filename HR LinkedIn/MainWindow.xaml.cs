using ExcelDataReader;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace HR_LinkedIn
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        DataTableCollection dataTableCollection;
        IWebDriver driver;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                textBox.Text= openFileDialog.FileName;
                using (var stream = File.Open(openFileDialog.FileName,FileMode.Open,FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true}
                        });
                        dataTableCollection = result.Tables;
                    }
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            driver = new ChromeDriver(Directory.GetCurrentDirectory());
            driver.Url = "https://www.linkedin.com/login";

            //Email
            var emailElement = driver.FindElement(By.Name("session_key"));
            emailElement.SendKeys("eugenia.jaramillo@prodigious.com");

            //Password
            var passwordElement = driver.FindElement(By.Name("session_password"));
            passwordElement.SendKeys("");

            var logInElement = driver.FindElement(By.ClassName("login__form_action_container"));
            logInElement.Click();


            DataTable table = dataTableCollection[0];


            bool anyRows = true;
            var i = 0;
            while (anyRows)
            {
                try
                {
                    DataRow row = table.Rows[i];
                    string cell = row[2].ToString();
                    if (!(cell == null || cell == ""))
                    {
                        ArrayList tabs = new ArrayList(driver.WindowHandles);

                        //Use the list of window handles to switch between windows
                        driver.SwitchTo().Window(tabs[tabs.Count - 1].ToString());

                        driver.Url = cell;
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);

                        var recruiterButton = driver.FindElement(By.ClassName("pv-s-profile-actions--view-profile-in-recruiter"));
                        recruiterButton.Click();

                        //Get the list of window handles
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                        tabs = new ArrayList(driver.WindowHandles);

                        //Use the list of window handles to switch between windows
                        driver.SwitchTo().Window(tabs[tabs.Count - 1].ToString());


                        try
                        {
                            var profileButton = driver.FindElement(By.ClassName("contract-list__item-buttons"));
                            profileButton.Click();
                        }
                        catch (Exception)
                        {


                        }


                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(4);

                        var inmailButton = driver.FindElement(By.ClassName("send-inmail-split-button"));
                        inmailButton.Click();

                        var subjectElement = driver.FindElement(By.Name("subject"));
                        subjectElement.SendKeys(textBoxSubject.Text);

                        var messageElement = driver.FindElement(By.ClassName("compose-txtarea"));
                        messageElement.SendKeys(textBoxMessage.Text);

                        var sendButton = driver.FindElement(By.ClassName("inmail-send-btn"));
                        sendButton.Click();
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(4);
                        driver.Close();
                        i++;
                    }
                    else
                    {
                        anyRows = false;
                    }
                }
                catch (Exception)
                {
                    anyRows = false;
                }
                
                
            }
            

            

        }
    }
}
