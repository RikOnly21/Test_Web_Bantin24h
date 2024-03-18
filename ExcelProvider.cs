using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestPart1
{
    public static class ExcelProvider
    {
        public static List<string> GetTestCaseDataFromExcel(int x)
        {

            List<string> data = new List<string>();

            using (var stream = File.Open(@"E:/School/BDCLPM/Testcase.xlsx", FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                Encoding encoding = Encoding.GetEncoding(1252);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[0];

                    var pos = table.Rows[x][0].ToString();
                    var _all = table.Rows[x][4].ToString();
                    var expected = table.Rows[x][5].ToString();
                    data.Add(pos);//vị trí của testcase
                    data.Add(_all);//datatest, có thể null
                    data.Add(expected);//kết quả mong đợi
                }
            }
            return data;
        }
        public static void WriteDataToExcel(string expected, string webrs, bool rs, string pos)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string filePath = @"E:/School/BDCLPM/Testcase.xlsx";

            try
            {
                FileInfo existingFile = new FileInfo(filePath);
                // Create a new Excel package
                using (var excelPackage = new ExcelPackage(existingFile))
                {
                    // Add a new worksheet
                    var worksheet = excelPackage.Workbook.Worksheets["Bình luận, yêu thích, chia sẻ"];

                    for (int i = 1; i < 100; i++)
                    {
                        if (worksheet.Cells["A" + i.ToString()].Value != null)
                        {

                            if (worksheet.Cells["A" + i].Value.ToString() == pos)
                            {
                                string realCell = "G" + i;
                                string rsCell = "H" + i;
                                if (rs)
                                {
                                    worksheet.Cells[realCell].Value = webrs;
                                    worksheet.Cells[rsCell].Value = "Passed";
                                }
                                else
                                {
                                    worksheet.Cells[realCell].Value = expected;
                                    worksheet.Cells[rsCell].Value = "Failed";
                                }



                            }
                        }
                    }

                    var fileInfo = new System.IO.FileInfo(filePath);
                    excelPackage.SaveAs(fileInfo);
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
        }
    }
}
