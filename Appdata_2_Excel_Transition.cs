using System.Xml.Linq;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Appdata_2_Excel_Transition
    {
        //static void Main(string[] args)
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nAppdata生成Excel - Transition");
            Console.ResetColor();

            Console.WriteLine("请输入XML文件路径：");
            string xmlFile = Console.ReadLine();

            // 去除路径空格自动双引号
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
                var worksheet = workbook.Workbook.Worksheets.Add("Sheet1");

                // 写入表头
                worksheet.Cells[1, 1].Value = "StateMachine";
                worksheet.Cells[1, 2].Value = "Transition src";
                worksheet.Cells[1, 3].Value = "Transition dest";
                worksheet.Cells[1, 4].Value = "Condition type";
                worksheet.Cells[1, 5].Value = "Condition path";
                worksheet.Cells[1, 5].Value = "Condition valvalue";

                int row = 2; // 行号

                // 解析XML文件
                XDocument doc = XDocument.Load(xmlFile);

                // 遍历所有的StateMachine元素
                foreach (var stateMachine in doc.Descendants().Where(x => string.Equals(x.Name.LocalName, "StateMachine", StringComparison.OrdinalIgnoreCase)))
                {
                    string name = stateMachine.Attribute("name")?.Value;

                    if (!string.IsNullOrEmpty(name))
                    {
                        // 输出StateMachine的name属性到第一列
                        worksheet.Cells[row, 1].Value = name;

                        // 获取Transition元素
                        var transitions = stateMachine.Descendants().Where(x => string.Equals(x.Name.LocalName, "Transition", StringComparison.OrdinalIgnoreCase));

                        // 遍历Transition元素
                        foreach (var transition in transitions)
                        {

                            //输出Transition src属性值到第二列
                            string transition_src = transition.Attribute("src")?.Value;

                            if (!string.IsNullOrEmpty(transition_src))
                            {
                                worksheet.Cells[row, 2].Value = transition_src;
                            }

                            //输出Transition dest属性值到第三列
                            string transition_dest = transition.Attribute("dest")?.Value;

                            if (!string.IsNullOrEmpty(transition_dest))
                            {
                                worksheet.Cells[row, 3].Value = transition_dest;
                            }

                            // 获取Condition元素
                            var conditions = transition.Descendants().Where(x => string.Equals(x.Name.LocalName, "Condition", StringComparison.OrdinalIgnoreCase));

                            int conditionRow = row; // 记录当前行号

                            // 遍历Condition元素，输出每个Condition的type属性值、path属性值和valvalue属性值到单独一行
                            foreach (var condition in conditions)
                            {
                                string conditionType = condition.Attribute("type")?.Value;
                                string conditionPath = condition.Attribute("path")?.Value;
                                string conditionValue = condition.Attribute("valvalue")?.Value;

                                if (!string.IsNullOrEmpty(conditionType))
                                {
                                    worksheet.Cells[conditionRow, 4].Value = conditionType;
                                    worksheet.Cells[conditionRow, 5].Value = conditionPath;
                                    worksheet.Cells[conditionRow, 6].Value = conditionValue;
                                    conditionRow++;
                                }
                            }

                            row = Math.Max(row, conditionRow); // 更新行号为最大值

                            // 添加空行分隔各个State
                            //row++;
                        }
                    }
                    else
                    {
                        // 如果StateMachine没有name属性，则跳过
                        continue;
                    }
                }

                // 自动调整列宽
                worksheet.Cells.AutoFitColumns();
                // 保存xlsx文件
                workbook.SaveAs(new FileInfo(xlsxFile));
            }
        }

        static string GetOutputFilePath(string xmlFile)
        {
            string directory = Path.GetDirectoryName(xmlFile);
            string fileName = Path.GetFileNameWithoutExtension(xmlFile);
            string xlsxFile = Path.Combine(directory, $"{fileName}_Transition.xlsx");
            return xlsxFile;
        }

    }
}