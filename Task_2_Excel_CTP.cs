using System.Xml.Linq;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Task_2_Excel_CTP
    {
        //static void Main(string[] args)
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nTask生成Excel - CTP");
            Console.ResetColor();

            Console.WriteLine("请输入XML文件路径：");
            string xmlFile = Console.ReadLine();

            //去除路径空格自动双引号
            xmlFile = xmlFile.Replace("\"", "");

            if (!File.Exists(xmlFile))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("指定的XML文件不存在\n");
                return;
            }

            string xlsxFile = GetOutputFilePath(xmlFile);

            ReadXml(xmlFile, xlsxFile);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("XLSX文件已生成：" + xlsxFile + "\n\n");

            //Console.WriteLine("按任意键退出...");
            //Console.ReadKey();
        }

        static void ReadXml(string xmlFile, string xlsxFile)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var workbook = new ExcelPackage())
            {

                var worksheet = workbook.Workbook.Worksheets.Add("Task");

                // 写入表头
                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Content";
                worksheet.Cells[1, 3].Value = "Tip";
                worksheet.Cells[1, 4].Value = "Paramkey";
                worksheet.Cells[1, 5].Value = "Errorhint";
                worksheet.Cells[1, 6].Value = "Waitfinishtip";

                int row = 2; // 行号

                // 解析XML文件
                XDocument doc = XDocument.Load(xmlFile);

                // 遍历所有的Task元素
                foreach (var task in doc.Descendants().Where(x => string.Equals(x.Name.LocalName, "Task", StringComparison.OrdinalIgnoreCase)))
                //foreach (var task in doc.Descendants("Task").Concat(doc.Descendants("task")))
                {
                    string name = task.Attribute("name")?.Value;
                    string content = task.Attribute("content")?.Value;
                    string tip = task.Attribute("tip")?.Value;
                    string errorhint = task.Attribute("errorhint")?.Value;
                    string waitfinishtip = task.Attribute("waitfinishtip")?.Value;

                    if (!string.IsNullOrEmpty(name))
                    {
                        // 输出Task属性到对应位置
                        worksheet.Cells[row, 1].Value = name;
                        worksheet.Cells[row, 2].Value = content;
                        worksheet.Cells[row, 3].Value = tip;
                        worksheet.Cells[row, 5].Value = errorhint;
                        worksheet.Cells[row, 6].Value = waitfinishtip;

                        // 获取Action元素
                        var action = task.Descendants().Where(x => string.Equals(x.Name.LocalName, "Action", StringComparison.OrdinalIgnoreCase));
                        //var action = task.Descendants("Action").Concat(task.Descendants("action"));


                        // 检查是否存在paramkey属性并输出到第四列
                        /*var actionParamkey = action.Where(x => !string.IsNullOrEmpty(x.Attribute("paramkey")?.Value)).Select(x => x.Attribute("paramkey")?.Value).FirstOrDefault();

                        if (!string.IsNullOrEmpty(actionParamkey))
                        {
                            worksheet.Cells[row, 4].Value = actionParamkey;
                            row++;
                        }
                        else
                        {
                            row++;
                        }*/

                        // 遍历Action元素，输出paramkey属性值到第四列
                        foreach (var state in action)
                        {
                            string actionParamkey = state.Attribute("paramkey")?.Value;

                            if (!string.IsNullOrEmpty(actionParamkey))
                            {
                                worksheet.Cells[row, 4].Value = actionParamkey;
                                row++;
                            }
                            else
                            {
                                row++;
                            }
                        }

                    }
                    else
                    {
                        // 如果task没有name属性，则跳过
                        continue;
                    }
                }

                // 去除从第1行到row行之间的所有空白行
                for (int i = 2; i <= row; i++)
                {
                    bool isRowEmpty = true;
                    for (int j = 1; j <= 6; j++)
                    {
                        if (worksheet.Cells[i, j].Value != null)
                        {
                            isRowEmpty = false;
                            break;
                        }
                    }

                    if (isRowEmpty)
                    {
                        worksheet.DeleteRow(i);
                        row--;
                        i--;
                    }
                }

                for (int i = 2; i < row; i++)
                {
                    worksheet.Cells[i, 4].Value = worksheet.Cells[i + 1, 4].Value;
                }


                // 去除从第1行到row行之间的所有空白行
                for (int i = 2; i <= row; i++)
                {
                    bool isRowEmpty = true;
                    for (int j = 1; j <= 6; j++)
                    {
                        if (worksheet.Cells[i, j].Value != null)
                        {
                            isRowEmpty = false;
                            break;
                        }
                    }

                    if (isRowEmpty)
                    {
                        worksheet.DeleteRow(i);
                        row--;
                        i--;
                    }
                }

                // 清空第row的数据
                worksheet.Cells[row, 4].Value = null;



                // 自动调整列宽
                worksheet.Cells.AutoFitColumns();
                // 冻结首行
                worksheet.View.FreezePanes(2, 1);
                // 保存xlsx文件
                workbook.SaveAs(new FileInfo(xlsxFile));
            }
        }

        static string GetOutputFilePath(string xmlFile)
        {
            string directory = Path.GetDirectoryName(xmlFile);
            string fileName = Path.GetFileNameWithoutExtension(xmlFile);
            string xlsxFile = Path.Combine(directory, $"{fileName}_CTP.xlsx");
            return xlsxFile;
        }
    }
}