
using ExcelDataReader;
using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;




using System.ComponentModel;
using System.Text;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net;


namespace TestPart1
{
    public class TestLove
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
        public void TestLov_01()
        {
            //truy cập web
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("start-maximized");
            IWebDriver driver = new ChromeDriver(@"C:/Users/Rik/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe", options);
            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);
            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(16);
            //kiểm tra 
            bool rs = false;
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("button")).Click();
            Thread.Sleep(1000);

            string webrs = driver.FindElement(By.XPath("(//div[@role='status'])[1]")).Text;

            if (webrs == lst[2])
            {
                rs = true;
            }
            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }

        [TestCase]
        public void TestLov_02()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);
            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(17);
            //kiểm tra 
            bool rs = false;
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("a")).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));//thời gian chờ 60s
            //Chờ load trang web
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/main/div/div/section[1]/div[4]/div[1]/button")));
            element.Click();
            Thread.Sleep(1500);
            string webrs = driver.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]")).Text;
            Thread.Sleep(10000);

            Console.WriteLine("Kết quả lỗi từ web:" + webrs);

            if (webrs == lst[2])
            {
                rs = true;
            }
            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }

        [TestCase]
        public void TestLov_03()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");

            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(18);
            string[] dataExcel = lst[1].Split(",");
            //kiểm tra 
            bool rs = false;

            IWebElement element = driver.FindElement(By.XPath("/html/body/div[1]/main/section[2]/div/div[1]/button"));
            string webrs1 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine(dataExcel[0]);
            Console.WriteLine(dataExcel[1]);
            Console.WriteLine(" ch nhấn " + webrs1);
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("button")).Click();
            Thread.Sleep(3000);
            string webrs2 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine("nhấn rồi " + webrs2);

            if ((webrs1.Trim() == dataExcel[0]) && (webrs2.Trim() == dataExcel[1]))
            {
                rs = true;
            }

            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }

        [TestCase]
        public void TestLov_04()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");

            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(19);
            string[] dataExcel = lst[1].Split(",");
            //kiểm tra 
            bool rs = false;
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("a")).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));//thời gian chờ 60s
            //Chờ load trang web
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/main/div/div/section[1]/div[4]/div[1]/button")));
            string webrs1 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine(dataExcel[0]);
            Console.WriteLine(dataExcel[1]);
            Console.WriteLine(" ch nhấn " + webrs1);
            element.Click();
            Thread.Sleep(3000);
            string webrs2 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine("nhấn rồi " + webrs2);

            if ((webrs1.Trim() == dataExcel[0]) && (webrs2.Trim() == dataExcel[1]))
            {
                rs = true;
            }

            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }
        [TestCase]
        public void TestLov_05()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");

            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);
            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(20);
            string[] dataExcel = lst[1].Split(",");
            //kiểm tra 
            bool rs = false;
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("a")).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));//thời gian chờ 60s
            //Chờ load trang web
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/main/div/div/section[1]/div[4]/div[1]/button")));
            string webrs1 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine(dataExcel[0]);
            Console.WriteLine(dataExcel[1]);
            Console.WriteLine(" ch nhấn " + webrs1.Trim());
            element.Click();
            Thread.Sleep(3000);
            string webrs2 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine("nhấn rồi " + webrs2.Trim());
            Console.WriteLine((webrs1.Trim() == dataExcel[0]));
            Console.WriteLine((webrs2.Trim() == dataExcel[1]));
            if ((webrs1.Trim() == dataExcel[0]) && (webrs2.Trim() == dataExcel[1]))
            {
                rs = true;
            }

            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }
        [TestCase]
        public void TestLov_06()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");

            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);
            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(21);
            string[] dataExcel = lst[1].Split(",");
            //kiểm tra 
            bool rs = false;
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("a")).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));//thời gian chờ 60s
            //Chờ load trang web
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/main/div/div/section[1]/div[4]/div[1]/button")));
            string webrs1 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine(dataExcel[0]);
            Console.WriteLine(dataExcel[1]);
            Console.WriteLine(" ch nhấn " + webrs1);
            element.Click();
            Thread.Sleep(3000);
            string webrs2 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine("nhấn rồi " + webrs2);

            if ((webrs1.Trim() == dataExcel[0]) && (webrs2.Trim() == dataExcel[1]))
            {
                rs = true;
            }

            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }

        [TestCase]
        public void TestLov_07()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");

            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);


            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(22);

            //Thực hiện test
            bool rs = false;
            //Ấn vào button yêu thích 
            IWebElement element = driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("button"));
            element.Click();
            Thread.Sleep(3000);
            //lưu lại trang thái khi đã nhấn vào button đó
            string webrs1 = element.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine("Trước khi đăng xuất : " + webrs1);
            //đăng xuất
            driver.FindElement(By.Id("react-aria-:Ricpla:")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[4]/div/div[1]/div/div/ul/li[4]/ul/li")).FindElement(By.TagName("span")).FindElement(By.TagName("button")).Click();
            Thread.Sleep(2000);
            driver.Navigate().Refresh();
            //Đăng nhập lại
            AutoLogin.AutoLog(driver);

            Thread.Sleep(2000);
            IWebElement element2 = driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("button"));
            string webrs2 = element2.FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
            Console.WriteLine("Sau khi đăng nhập lại : " + webrs2);
            //Kiểm tra
            if (webrs1 == webrs2)
            {
                rs = true;
            }

            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }
        [TestCase]
        public void TestLov_08()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");

            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);
            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(23);
            string[] dataExcel = lst[1].Split(",");
            //Thực hiện test
            bool rs = true;
            //Thực hiện tìm kiếm
            driver.FindElement(By.TagName("input")).Click();
            driver.FindElement(By.TagName("input")).SendKeys(dataExcel[0]);
            Thread.Sleep(500);
            driver.FindElement(By.TagName("button")).Click();
            Thread.Sleep(3000);
            //trỏ đến kết quả tìm kiếm
            IWebElement section1 = driver.FindElement(By.TagName("section")).FindElements(By.TagName("section"))[1];
            foreach (var a in section1.FindElements(By.TagName("section")))
            {
                a.FindElement(By.TagName("button")).Click();
                Thread.Sleep(1000);
                string webrs = a.FindElement(By.TagName("button")).FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
                Console.WriteLine(a.FindElement(By.TagName("button")).FindElement(By.TagName("svg")).GetAttribute("Class").ToString());
                if (dataExcel[1] != webrs.Trim())
                {
                    rs = false;
                }
            }
            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }
        [TestCase]
        public void TestLov_09()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");

            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);


            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(24);
            string[] dataExcel = lst[1].Split(",");
            //Thực hiện test
            bool rs = true;
            //Thực hiện tìm kiếm
            driver.FindElement(By.TagName("input")).Click();
            driver.FindElement(By.TagName("input")).SendKeys(dataExcel[0]);
            Thread.Sleep(500);
            driver.FindElement(By.TagName("button")).Click();
            Thread.Sleep(3000);
            //trỏ đến kết quả tìm kiếm
            IWebElement section1 = driver.FindElement(By.TagName("section")).FindElements(By.TagName("section"))[1];
            foreach (var a in section1.FindElements(By.TagName("section")))
            {
                a.FindElement(By.TagName("button")).Click();
                Thread.Sleep(1000);
                string webrs = a.FindElement(By.TagName("button")).FindElement(By.TagName("svg")).GetAttribute("Class").ToString();
                Console.WriteLine(a.FindElement(By.TagName("button")).FindElement(By.TagName("svg")).GetAttribute("Class").ToString());
                if (dataExcel[1] != webrs.Trim())//phải trim vì đôi khi kết quả web trả về có khoảng trống ở sau
                {
                    rs = false;
                }
            }
            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }
        [TestCase]
        public void TestLov_10()
        {
            //truy cập web
            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);

            //đăng nhập
            AutoLogin.AutoLog(driver);


            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(25);

            Thread.Sleep(3000);

            //Thực hiện test
            bool rs = false;
            string webrs1 = driver.FindElement(By.XPath("/html/body/div[1]/main/section[2]/div/div[1]/div[2]/div/a/h2")).Text.Trim();
            driver.FindElements(By.TagName("section"))[1].FindElement(By.TagName("div")).FindElement(By.TagName("div")).FindElement(By.TagName("button")).Click();
            Console.WriteLine("Bản tin thực hiện yêu thích " + webrs1);
            driver.FindElement(By.Id("react-aria-:Ricpla:")).Click();
            Thread.Sleep(1000);

            //trỏ vào phần tử bằng cách tìm text
            IWebElement element = driver.FindElement(By.XPath("//a[text()='Tin yêu thích']"));
            element.Click();
            Thread.Sleep(2000);
            driver.Navigate().Refresh();
            Thread.Sleep(2000);
            string webrs2 = driver.FindElement(By.XPath("/html/body/div[1]/main/div/div[2]/div/div[2]/div[1]/a/div[2]/h3")).Text.Trim();
            Console.WriteLine("Bản tin yêu thích đầu tiên:" + webrs2);
            if (webrs1 == webrs2)
            {
                rs = true;
            }
            //ghi dữ liệu vào excel expected, webrs, rs, position

            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }
        [TestCase]
        public void TestLov_11()
        {
            //truy cập web
            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);
            AutoLogin.AutoLog(driver);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(26);

            //Thực hiện test
            bool rs = false;
            driver.FindElement(By.Id("react-aria-:Ricpla:")).Click();
            Thread.Sleep(1000);

            //trỏ vào phần tử bằng cách tìm text
            IWebElement element = driver.FindElement(By.XPath("//a[text()='Tin yêu thích']"));
            element.Click();
            Thread.Sleep(2000);
            driver.Navigate().Refresh();
            Thread.Sleep(2000);
            string webrs1 = driver.FindElement(By.XPath("/html/body/div[1]/main/div/div[2]/div/div[2]/div[1]/a/div[2]/h3")).Text.Trim();
            Console.WriteLine("Bản tin yêu thích đầu tiên:" + webrs1);

            //Ấn nút hủy yêu thích
            driver.FindElement(By.XPath("/html/body/div[1]/main/div/div[2]/div/div[2]/div[1]/section/div/div[2]/button[1]")).Click();
            Thread.Sleep(2000);
            string webrs2 = driver.FindElement(By.XPath("/html/body/div[1]/main/div/div[2]/div/div[2]/div[1]/a/div[2]/h3")).Text.Trim();
            Console.WriteLine("Bản tin yêu thích đầu tiên sau khi hủy yêu thích:" + webrs2);
            Thread.Sleep(2000);

            if (webrs1 != webrs2)
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