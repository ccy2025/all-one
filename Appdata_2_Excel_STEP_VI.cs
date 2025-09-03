using System.Xml;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Appdata_2_Excel_STEP_VI
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nAppdata生成Excel - HS-STEP_VI - StateMachine");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("仅支持标准HS-STEP_VI-AppData互转");
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

            // 解析XML文件
            //XDocument doc = XDocument.Load(xmlFile);

            // 加载XML文件
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFile);

            using (var workbook = new ExcelPackage())
            {
                var worksheet = workbook.Workbook.Worksheets.Add("STEP-VI");

                // 写入表头
                worksheet.Cells[1, 1].Value = "MOD";
                worksheet.Cells[1, 2].Value = "NAME";
                worksheet.Cells[1, 3].Value = "highlight";
                worksheet.Cells[1, 4].Value = "setlist";
                worksheet.Cells[1, 5].Value = "stop_oVar";
                worksheet.Cells[1, 6].Value = "stop_valvalue";

                string col1Value = doc.SelectSingleNode("/AppData/CustomStateMachine/StateMachine/StateMachine/Transition[@src='normal']/Condition/Condition[@type='oVar']")?.Attributes["path"]?.Value?.Split('_')[0] ?? "NULL";
                string col2Value = doc.SelectSingleNode("/AppData/CustomStateMachine/StateMachine/StateMachine/Transition[@src='normal']/Condition/Condition[@type='oVar']")?.Attributes["path"]?.Value?.Split('_')[1] ?? "NULL";

                worksheet.Cells[2, 1].Value = col1Value;
                worksheet.Cells[2, 2].Value = col2Value;

                // 获取StateMachine节点
                XmlNodeList stateMachineNode = doc.SelectNodes("/AppData/CustomStateMachine/StateMachine/StateMachine[Transition/Condition/Condition[@type='operationmode' and @modeid='0']]");

                // 定义行索引
                int rowIndex = 2;

                // 遍历State节点
                foreach (XmlNode stateNode in stateMachineNode)
                {
                    string col3Value = stateNode.SelectSingleNode("State/Action/Action[@type='highlight']")?.Attributes["path"]?.Value ?? "NULL";
                    string col4Value = stateNode.SelectSingleNode("State/Action/Action[@type='setlist']")?.Attributes["value"]?.Value ?? "NULL";
                    string col5Value = stateNode.SelectSingleNode("Transition[@dest='stop']/Condition/Condition[@type='oVar']")?.Attributes["path"]?.Value?.Split('_')[1] ?? "NULL";
                    string col6Value = stateNode.SelectSingleNode("Transition[@dest='stop']/Condition/Condition[@type='oVar']")?.Attributes["valvalue"]?.Value ?? "NULL";

                    worksheet.Cells[rowIndex, 3].Value = col3Value;
                    worksheet.Cells[rowIndex, 4].Value = col4Value;
                    worksheet.Cells[rowIndex, 5].Value = col5Value;
                    worksheet.Cells[rowIndex, 6].Value = col6Value;

                    rowIndex++;
                }

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
            string xlsxFile = Path.Combine(directory, $"{fileName}_STEP-VI.xlsx");
            return xlsxFile;
        }
    }
}