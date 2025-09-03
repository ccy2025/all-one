using System.Xml;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Excel_2_Appdata_Device
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nExcel生成Appdata - Device");
            Console.WriteLine("(XmlElement)");
            Console.ResetColor();

            Console.WriteLine("请输入Excel文件路径：");
            string excelFilePath = Console.ReadLine();

            // 去除路径空格自动双引号
            excelFilePath = excelFilePath.Replace("\"", "");

            if (!File.Exists(excelFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("指定的Excel文件不存在\n");
                return;
            }

            // 获取 Excel 文件所在的目录
            string directoryPath = Path.GetDirectoryName(excelFilePath);

            // 创建与 Excel 文件名称相同的 XML 文件
            string xmlFilePath = Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(excelFilePath) + "_Device.xml");

            // 打开 Excel 文件
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // 读取第一个工作表
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // 创建 XML 文档
                XmlDocument xmlDoc = new XmlDocument();
                XmlElement rootElement = xmlDoc.CreateElement("AppData");
                xmlDoc.AppendChild(rootElement);

                //int rowCount = worksheet.Dimension.Rows;

                // 遍历数据，将其添加到 XML 中
                for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
                {
                    XmlElement deviceElement = xmlDoc.CreateElement("Device");

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string headerText = worksheet.Cells[2, col].Text;
                        string cellText = worksheet.Cells[row, col].Text;

                        if (!string.IsNullOrWhiteSpace(cellText))
                        {
                            XmlElement element = xmlDoc.CreateElement(headerText);
                            element.InnerText = cellText;
                            deviceElement.AppendChild(element);
                        }
                    }

                    rootElement.AppendChild(deviceElement);

                    // 在每个device元素后添加两个换行符，以确保有一个空行的间隔
                    //XmlText newline = xmlDoc.CreateTextNode("\n");
                    //rootElement.AppendChild(newline);

                }


                // 保存 XML 文件
                xmlDoc.Save(xmlFilePath);

                //Console.WriteLine($"XML 文件已生成： {xmlFilePath}");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("XML文件已生成：" + xmlFilePath + "\n\n");
            }

            //Console.WriteLine("按任意键退出...");
            //Console.ReadKey();

        }

    }
}