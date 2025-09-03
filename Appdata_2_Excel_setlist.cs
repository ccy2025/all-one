using System.Xml.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace All_in_1
{
    internal class Appdata_2_Excel_setlist
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nAppdata生成Excel - setlist");
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
                var worksheet = workbook.Workbook.Worksheets.Add("setlist");

                // 写入表头
                worksheet.Cells[1, 1].Value = "StateMachine";
                worksheet.Cells[1, 2].Value = "State";
                worksheet.Cells[1, 3].Value = "setlist value";

                int row = 2; // 行号

                // 解析XML文件
                XDocument doc = XDocument.Load(xmlFile);

                // 对Action type setlist 大小写不敏感，避免不规范Appdata
                // 影响性能
                var actions = doc.Descendants()
                    .Where(e => string.Equals(e.Name.LocalName, "action", StringComparison.OrdinalIgnoreCase))
                    .Where(a => a.Attributes()
                                 .FirstOrDefault(attr => string.Equals(attr.Name.LocalName, "type", StringComparison.OrdinalIgnoreCase))?.Value
                                 .Equals("setlist", StringComparison.OrdinalIgnoreCase) == true);

                foreach (var action in actions)
                {
                    // 获取父级State节点的name属性
                    //var stateName = action.Parent.Parent.Attribute("name").Value;
                    // name大小写不敏感处理
                    var stateName = action.Parent.Parent.Attributes()
                        .FirstOrDefault(attr => string.Equals(attr.Name.LocalName, "name", StringComparison.OrdinalIgnoreCase))?.Value;

                    // 获取祖父级Statemachine节点的name属性
                    //var statemachineName = action.Parent.Parent.Parent.Attribute("name").Value;
                    var statemachineName = action.Parent.Parent.Parent.Attributes()
                        .FirstOrDefault(attr => string.Equals(attr.Name.LocalName, "name", StringComparison.OrdinalIgnoreCase))?.Value;

                    // 获取Action节点的value属性
                    //var actionValue = action.Attribute("value").Value;
                    var actionValue = action.Attributes()
                        .FirstOrDefault(attr => string.Equals(attr.Name.LocalName, "value", StringComparison.OrdinalIgnoreCase))?.Value;

                    // 数据写入Excel
                    worksheet.Cells[row, 1].Value = statemachineName;
                    worksheet.Cells[row, 2].Value = stateName;
                    worksheet.Cells[row, 3].Value = actionValue;

                    row++;
                }

                // 自动调整列宽
                worksheet.Cells.AutoFitColumns();

                //合并第一列相同单元格
                object previousCellValue = null;
                int startRow = 2;

                for (int checkRow = 1; checkRow <= worksheet.Dimension.End.Row; checkRow++)
                {
                    var currentCellValue = worksheet.Cells[checkRow, 1].Value; // 只检查第一列
                    if (currentCellValue != null && currentCellValue.Equals(previousCellValue))
                    {
                        // 当前单元格与上一个单元格内容相同，则继续
                        continue;
                    }
                    else
                    {
                        // 如果内容不同，则检查是否有需要合并的单元格
                        if (checkRow - startRow > 1) // 说明之前有连续相同的单元格需要合并
                        {
                            var mergeRange = worksheet.Cells[startRow, 1, checkRow - 1, 1];
                            mergeRange.Merge = true;
                            // 设置垂直对齐方式为居中
                            mergeRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        startRow = checkRow; // 重置起始行为当前行
                    }
                    previousCellValue = currentCellValue;
                }

                // 循环结束后，检查最后一段是否需要合并
                if (worksheet.Dimension.End.Row - startRow > 0)
                {
                    var mergeRange = worksheet.Cells[startRow, 1, worksheet.Dimension.End.Row, 1];
                    mergeRange.Merge = true;
                    // 设置垂直对齐方式为居中
                    mergeRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                // 自动调整列宽
                //worksheet.Cells.AutoFitColumns();
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
            string xlsxFile = Path.Combine(directory, $"{fileName}_setlist.xlsx");
            return xlsxFile;
        }
    }
}