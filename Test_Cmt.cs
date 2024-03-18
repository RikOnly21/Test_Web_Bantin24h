using DocumentFormat.OpenXml.Bibliography;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace TestPart1
{
    internal class Test_Cmt
    {

        private IWebDriver driver;
        [SetUp]
        public void SetUp()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("start-maximized");
            options.AddArgument("--enable-features=SameSiteByDefaultCookies,CookiesWithoutSameSiteMustBeSecure");
            options.AddAdditionalOption("useAutomationExtension", false);
            options.AddUserProfilePreference("profile.default_content_setting_values.cookies", 1);
            options.AddUserProfilePreference("profile.cookie_controls_mode", 0);
            driver = new ChromeDriver(@"C:/Users/Rik/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe", options);
        }

        [TestCase]
        public void Test_Cmt01()
        {
            //truy cập web

            driver.Navigate().GoToUrl("http://192.168.1.7:3000");
            Thread.Sleep(3000);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(28);

            //Thực hiện test
            bool rs = false;
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("a")).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));//thời gian chờ 60s
            //Chờ load trang web
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/main/div/div/section[1]/div[2]")));

            IList<IWebElement> elements = driver.FindElements(By.TagName("textarea"));
            if (elements.Count == 0)
            {
                rs = true;
            }

            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }


        [TearDown]
        public void TearDown()
        {
            driver?.Quit();
            driver?.Dispose();
        }
    }
}
