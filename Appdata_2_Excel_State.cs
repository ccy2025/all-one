using System.Xml.Linq;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Appdata_2_Excel_State
    {
        //static void Main(string[] args)
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nAppdata生成Excel - State");
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
                var worksheet = workbook.Workbook.Worksheets.Add("Sheet1");

                // 写入表头
                worksheet.Cells[1, 1].Value = "StateMachine";
                worksheet.Cells[1, 2].Value = "State";
                worksheet.Cells[1, 3].Value = "Action Type";
                worksheet.Cells[1, 4].Value = "Action Path";
                worksheet.Cells[1, 5].Value = "Active Value";

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

                        // 获取State元素
                        var states = stateMachine.Descendants().Where(x => string.Equals(x.Name.LocalName, "State", StringComparison.OrdinalIgnoreCase));

                        // 遍历State元素，输出name属性值到第二列
                        foreach (var state in states)
                        {
                            string stateName = state.Attribute("name")?.Value;

                            if (!string.IsNullOrEmpty(stateName))
                            {
                                worksheet.Cells[row, 2].Value = stateName;
                            }

                            // 获取Action元素
                            var actions = state.Descendants().Where(x => string.Equals(x.Name.LocalName, "Action", StringComparison.OrdinalIgnoreCase));

                            int actionRow = row; // 记录当前行号

                            // 遍历Action元素，输出每个Action的type属性值、path属性值和active属性值到单独一行
                            foreach (var action in actions)
                            {
                                string actionType = action.Attribute("type")?.Value;
                                string actionPath = action.Attribute("path")?.Value;
                                string activeValue = action.Attribute("value")?.Value;

                                if (!string.IsNullOrEmpty(actionType))
                                {
                                    worksheet.Cells[actionRow, 3].Value = actionType;
                                    worksheet.Cells[actionRow, 4].Value = actionPath;
                                    worksheet.Cells[actionRow, 5].Value = activeValue;
                                    actionRow++;
                                }
                            }

                            row = Math.Max(row, actionRow); // 更新行号为最大值

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
            string xlsxFile = Path.Combine(directory, $"{fileName}_State.xlsx");
            return xlsxFile;
        }
    }
}