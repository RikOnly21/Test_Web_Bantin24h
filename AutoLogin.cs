using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestPart1
{
    public static class AutoLogin
    {
        public static void AutoLog(IWebDriver driver)
        {
            driver.FindElement(By.Id("react-aria-:Ricpla:")).Click();
            Thread.Sleep(500);
            driver.FindElement(By.Id("react-aria-:RicplaH1:")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("#identifier-field")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("#identifier-field")).SendKeys("nhocngusi2003@gmail.com");
            driver.FindElement(By.CssSelector("button.cl-formButtonPrimary.🔒️.cl-button")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.Id("password-field")).Click();
            driver.FindElement(By.Id("password-field")).SendKeys("test123123");
            driver.FindElement(By.CssSelector("button.cl-formButtonPrimary.cl-button")).Click();

            Thread.Sleep(2000);
            driver.Navigate().Refresh();
        }
    }
}
