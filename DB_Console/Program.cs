using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            ChromeOptions option = new ChromeOptions();
            //option.AddArguments(@"user-data-dir=C:\Users\Dell\AppData\Local\Google\Chrome\User Data\Profile 1");
            //IWebDriver driver = new ChromeDriver(option);

            ChromeDriverService serviceBB = ChromeDriverService.CreateDefaultService(@"D:\SMS_Myanmarmese\RekutenPay\RekutenPay\bin\Debug\");
            serviceBB.HideCommandPromptWindow = false; //ams.12.4.18        
            option.AddArguments(@"user-data-dir=C:\Users\Dell\AppData\Local\Google\Chrome\User Data\Default");
            using (IWebDriver Firefox = new ChromeDriver(serviceBB))
            {
                
                Firefox.Url = "https://www.google.com";
                System.Threading.Thread.Sleep(100000);
            }

            }
    }
}
