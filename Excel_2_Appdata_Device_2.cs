using System.Text;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Excel_2_Appdata_Device_2
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nExcel生成Appdata - Device");
            Console.WriteLine("(StringBuilder)");
            Console.ResetColor();

            Console.WriteLine("请输入Excel文件路径：");
            string excelFilePath = Console.ReadLine();

            //去除路径空格自动双引号
            excelFilePath = excelFilePath.Replace("\"", "");

            if (!File.Exists(excelFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("指定的Excel文件不存在\n");
                return;
            }

            // 获取 Excel 文件所在的目录
            string directoryPath = Path.GetDirectoryName(excelFilePath);

            // 创建一个与 Excel 文件名称相同的 XML 文件
            string xmlFilePath = Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(excelFilePath) + "_Device.xml");

            // 打开 Excel 文件
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // 读取第一个工作表
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // 手动方式构建 不够优雅 留着以后再改 2024.2.5
                // 构建 XML
                StringBuilder xmlStringBuilder = new StringBuilder();
                //xmlStringBuilder.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                xmlStringBuilder.AppendLine("<AppData>");

                for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
                {
                    xmlStringBuilder.AppendLine("\n\t<Device>");

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string headerText = worksheet.Cells[2, col].Text;
                        string cellText = worksheet.Cells[row, col].Text;

                        if (!string.IsNullOrWhiteSpace(cellText))
                        {
                            xmlStringBuilder.AppendLine($"\t\t<{headerText}>{System.Security.SecurityElement.Escape(cellText)}</{headerText}>");
                        }
                    }

                    xmlStringBuilder.AppendLine("\t</Device>");
                }

                xmlStringBuilder.AppendLine("\n</AppData>");

                // 将构建好的 XML 字符串写入文件
                File.WriteAllText(xmlFilePath, xmlStringBuilder.ToString(), Encoding.UTF8);

                //Console.WriteLine($"XML 文件已生成： {xmlFilePath}");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("XML文件已生成：" + xmlFilePath + "\n\n");
            }

            //Console.WriteLine("按任意键退出...");
            //Console.ReadKey();

        }

    }
}