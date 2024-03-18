using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestPart1
{
    public class Test_Gui
    {
        private IWebDriver driver;
        [SetUp]
        public void SetUp()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("start-maximized");
            driver = new ChromeDriver(@"C:/Users/Rik/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe", options);
        }

        [TestCase]
        public void TestGui_01()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(1000);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(9);

            //kiểm tra
            string webrs = driver.FindElement(By.TagName("body")).GetCssValue("font-family").ToString();
            webrs = webrs.Replace("\"", "");//loại bỏ dấu ngoặc 
            bool rs = (webrs == lst[2]);//gán tạm kết quả

            //Ghi dữ liệu vào file excel;
            ExcelProvider.WriteDataToExcel(lst[2], webrs, rs, lst[0]);
            driver.Close();
            Assert.That(rs, Is.True);

        }

        [TestCase]
        public void TestGui_02()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(10);
            //kiểm tra 
            string widthrs = "Width = " + driver.FindElement(By.TagName("button")).GetCssValue("width").ToString();
            string heightrs = "Height = " + driver.FindElement(By.TagName("button")).GetCssValue("height").ToString();

            string[] datatest = lst[2].Trim().Split(",");
            bool rs = false;
            if (widthrs == datatest[0] && heightrs == datatest[1])
            {
                rs = true;
            }

            //ghi dữ liệu vào excel
            ExcelProvider.WriteDataToExcel(lst[2], widthrs + heightrs, rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);

        }

        [TestCase]
        public void TestGui_03()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(11);
            //kiểm tra 
            Console.WriteLine("Số phần tử có tagName = svg là " + driver.FindElements(By.TagName("svg")).Count);
            bool rs = true;
            string widthrs = "";
            string heightrs = "";
            foreach (var a in driver.FindElements(By.TagName("svg")))//duyệt tất cả các phần tử dạng svg
            {
                widthrs = a.GetCssValue("width").ToString();
                heightrs = a.GetCssValue("height").ToString();
                if (widthrs != heightrs)
                {
                    rs = false;
                }
            }
            string datatest = lst[2];
            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);

        }


        [TestCase]
        public void TestGui_04()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);

            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected
            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(12);
            //Console.WriteLine(driver.FindElement(By.TagName("body")).GetCssValue("color-scheme"));
            //kiểm tra 
            bool rs = false;
            string webrs = "Color-scheme:" + driver.FindElement(By.TagName("body")).GetCssValue("color-scheme").ToString();
            if (webrs == lst[2])
            {
                rs = true;
            }

            //ghi dữ liệu vào excel
            ExcelProvider.WriteDataToExcel(lst[2], webrs, rs, lst[0]);

            //
            driver.Close();
            Assert.That(rs, Is.True);

        }

        [TestCase]
        public void TestGui_05()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);


            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected :3/7/2024

            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(13);
            //kiểm tra 
            bool rs = false;
            string dateRs = driver.FindElement(By.TagName("header")).FindElement(By.TagName("div")).FindElement(By.TagName("span")).Text;
            Console.WriteLine("kết quả từ web :" + dateRs);

            string[] dataExcel = lst[2].Split('/');
            dataExcel[2] = "" + dataExcel[2][0] + dataExcel[2][1] + dataExcel[2][2] + dataExcel[2][3];

            string expectedRs = dataExcel[1] + " tháng " + dataExcel[0] + ", " + dataExcel[2];
            Console.WriteLine("kết quả mong đợi :" + expectedRs);
            if (expectedRs == dateRs)
            { rs = true; }


            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], lst[2], rs, lst[0]);
            //
            driver.Close();
            Assert.That(rs, Is.True);
        }

        [TestCase]
        public void TestGui_06()
        {
            //truy cập web

            driver.Navigate().GoToUrl("https://bantin24h.asakuri.tech");
            Thread.Sleep(3000);


            //lấy dữ liệu từ excel 0 là pos, 1 là datatest, 2 là expected :3/7/2024

            List<string> lst = ExcelProvider.GetTestCaseDataFromExcel(14);
            //kiểm tra 
            bool rs = true;
            Console.WriteLine("Số phần tử có tagName = img là " + driver.FindElements(By.TagName("img")).Count);
            string widthrs = "";
            string heightrs = "";
            string webrs = "";
            foreach (var a in driver.FindElements(By.TagName("img")))//duyệt tất cả các phần tử dạng svg
            {
                widthrs = a.GetCssValue("width").ToString();
                heightrs = a.GetCssValue("height").ToString();
                if (widthrs != heightrs)
                {
                    Console.WriteLine("Các giá trị khác nhau là ");
                    Console.WriteLine("width :" + widthrs);
                    Console.WriteLine("height :" + heightrs);
                    rs = false;
                    webrs = "Width!=Height";

                }
            }
            //ghi dữ liệu vào excel expected, webrs, rs, position
            ExcelProvider.WriteDataToExcel(lst[2], webrs, rs, lst[0]);
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
