using System.Xml.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace All_in_1
{
    internal class Appdata_2_Excel_StateMachine
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nAppdata生成Excel - StateMachine");
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
                var worksheet = workbook.Workbook.Worksheets.Add("StateMachine");

                // 写入表头
                worksheet.Cells[1, 1].Value = "StateMachine Name";
                worksheet.Cells[1, 2].Value = "State Name";
                worksheet.Cells[1, 3].Value = "Action Type";
                //worksheet.Cells[1, 4].Value = "Action Path";
                //worksheet.Cells[1, 5].Value = "Active Value";
                worksheet.Cells[1, 4].Value = "Transition src";
                worksheet.Cells[1, 5].Value = "Transition dest";
                worksheet.Cells[1, 6].Value = "Condition Type";
                worksheet.Cells[1, 7].Value = "Condition Path";
                worksheet.Cells[1, 8].Value = "valvalue";

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
                        //worksheet.Cells[row, 1].Value = name;

                        // 获取State元素
                        var states = stateMachine.Descendants().Where(x => string.Equals(x.Name.LocalName, "State", StringComparison.OrdinalIgnoreCase));

                        // 遍历State元素，输出name属性值到第二列
                        foreach (var state in states)
                        {
                            string stateName = state.Attribute("name")?.Value;

                            //if (!string.IsNullOrEmpty(stateName))
                            //{
                            //    worksheet.Cells[row, 2].Value = stateName;
                            //}

                            // 获取Action元素
                            var actions = state.Descendants().Where(x => string.Equals(x.Name.LocalName, "Action", StringComparison.OrdinalIgnoreCase));
                            //var actions = state.Descendants("action").Concat(state.Descendants("Action"));

                            //int actionRow = row; // 记录当前行号

                            // 遍历Action元素，输出每个Action的type属性值、path属性值和active属性值到单独一行
                            foreach (var action in actions)
                            {
                                string actionType = action.Attribute("type")?.Value;
                                string actionPath = action.Attribute("path")?.Value;
                                //string activeValue = GetActiveValue(action, actionType);
                                string activeValue = action.Attribute("value")?.Value;

                                if (!string.IsNullOrEmpty(actionType) && actionType != "group")
                                //if (!string.IsNullOrEmpty(actionType) && actionType != "group")
                                {
                                    worksheet.Cells[row, 1].Value = name;
                                    worksheet.Cells[row, 2].Value = stateName;
                                    worksheet.Cells[row, 3].Value = actionType;
                                    //worksheet.Cells[row, 4].Value = actionPath;
                                    //worksheet.Cells[row, 5].Value = activeValue;
                                    row++;
                                }
                            }

                            //row = Math.Max(row, actionRow); // 更新行号为最大值

                            // 添加空行分隔各个State
                            //row++;
                        }

                        // 获取Transition元素
                        var transitions = stateMachine.Descendants().Where(x => string.Equals(x.Name.LocalName, "Transition", StringComparison.OrdinalIgnoreCase));
                        //var transitions = stateMachine.Descendants("Transition").Concat(stateMachine.Descendants("transition"));

                        // 遍历Transition元素
                        foreach (var transition in transitions)
                        {

                            //输出Transition src属性值到第六列
                            string transition_src = transition.Attribute("src")?.Value;

                            /*if (!string.IsNullOrEmpty(transition_src))
                            {
                                worksheet.Cells[row, 6].Value = transition_src;
                            }*/

                            //输出Transition dest属性值到第七列
                            string transition_dest = transition.Attribute("dest")?.Value;

                            /*if (!string.IsNullOrEmpty(transition_dest))
                            {
                                worksheet.Cells[row, 7].Value = transition_dest;
                            }*/

                            // 获取Condition元素
                            var conditions = transition.Descendants().Where(x => string.Equals(x.Name.LocalName, "Condition", StringComparison.OrdinalIgnoreCase));
                            //var conditions = transition.Descendants("Condition").Concat(transition.Descendants("condition"));

                            //int conditionRow = row; // 记录当前行号

                            // 遍历Condition元素，输出每个Condition的type属性值、path属性值和valvalue属性值到单独一行
                            foreach (var condition in conditions)
                            {
                                string conditionType = condition.Attribute("type")?.Value;
                                string conditionPath = condition.Attribute("path")?.Value;
                                string conditionValue = condition.Attribute("valvalue")?.Value;

                                if (!string.IsNullOrEmpty(conditionType) && conditionType != "and" && conditionType != "or")
                                {
                                    worksheet.Cells[row, 1].Value = name;
                                    //worksheet.Cells[row, 2].Value = stateName;

                                    worksheet.Cells[row, 4].Value = transition_src;
                                    worksheet.Cells[row, 5].Value = transition_dest;
                                    worksheet.Cells[row, 6].Value = conditionType;
                                    worksheet.Cells[row, 7].Value = conditionPath;
                                    worksheet.Cells[row, 8].Value = conditionValue;
                                    row++;

                                    //worksheet.Cells[conditionRow, 8].Value = conditionType;
                                    //worksheet.Cells[conditionRow, 9].Value = conditionPath;
                                    //worksheet.Cells[conditionRow, 10].Value = conditionValue;
                                    //conditionRow++;
                                }
                            }

                            //row = Math.Max(row, conditionRow); // 更新行号为最大值

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

                // 合并第1，2，4，5列相邻且相同单元格
                int[] columnsToMerge = { 1, 2, 4, 5 };

                foreach (int columnToMerge in columnsToMerge)
                {
                    object previousCellValue = null;
                    int startRow = 2;

                    for (int checkRow = 1; checkRow <= worksheet.Dimension.End.Row; checkRow++)
                    {
                        var currentCellValue = worksheet.Cells[checkRow, columnToMerge].Value;
                        if (currentCellValue != null && currentCellValue.Equals(previousCellValue))
                        {
                            continue;
                        }
                        else
                        {
                            if (checkRow - startRow > 1)
                            {
                                var mergeRange = worksheet.Cells[startRow, columnToMerge, checkRow - 1, columnToMerge];
                                mergeRange.Merge = true;
                                mergeRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                            startRow = checkRow;
                        }
                        previousCellValue = currentCellValue;
                    }

                    if (worksheet.Dimension.End.Row - startRow > 0)
                    {
                        var mergeRange = worksheet.Cells[startRow, columnToMerge, worksheet.Dimension.End.Row, columnToMerge];
                        mergeRange.Merge = true;
                        mergeRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                }

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
            string xlsxFile = Path.Combine(directory, $"{fileName}_StateMachine.xlsx");
            return xlsxFile;
        }

    }
}