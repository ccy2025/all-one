using System.Xml.Linq;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Paper_2_Excel
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nPaper生成Excel");
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

                var worksheet = workbook.Workbook.Worksheets.Add("Paper");

                // 写入标题行
                worksheet.Cells[1, 1].Value = "项目名称";
                worksheet.Cells[1, 2].Value = "过程";
                worksheet.Cells[1, 3].Value = "变量名";
                worksheet.Cells[1, 4].Value = "条件";
                worksheet.Cells[1, 5].Value = "变量值";
                worksheet.Cells[1, 6].Value = "题目分值";
                worksheet.Cells[1, 7].Value = "上偏差";
                worksheet.Cells[1, 8].Value = "下偏差";
                worksheet.Cells[1, 9].Value = "最大上偏差";
                worksheet.Cells[1, 10].Value = "最大下偏差";
                worksheet.Cells[1, 11].Value = "扣分上限";
                worksheet.Cells[1, 12].Value = "题目描述";

                // 解析XML文件
                XDocument doc = XDocument.Load(xmlFile);

                // 读取ProjectName属性值
                string projectName = doc.Element("AppData").Element("ProjectName").Value;
                worksheet.Cells[2, 1].Value = projectName;

                // 构建ProcedureID到Name的映射
                Dictionary<string, string> procedureIdToName = doc.Descendants("Procedures")
                    .ToDictionary(
                        p => p.Element("ID").Value,
                        p => p.Element("Name").Value
                    );

                // 没必要限制大小写
                //string? projectName = doc.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ID", StringComparison.OrdinalIgnoreCase))?.Value;

                int row = 2; // 开始填充数据的起始行

                // 遍历所有的Device元素
                // 遍历XML中的Questions节点
                foreach (XElement question in doc.Descendants("Questions"))
                {
                    // 忽略Procedures内部的Questions节点
                    if (question.Parent.Name != "AppData") continue;

                    // 读取ProcedureID
                    string procedureID = question.Element("ProcedureID")?.Value;

                    // 查找对应的Procedure Name
                    string procedureName = procedureIdToName.ContainsKey(procedureID) ? procedureIdToName[procedureID] : "N/A";

                    string? ScoreConditionExpression = question.Element("ScoreCondition")?.Element("Expression")?.Value;
                    string? ScoreConditionVerb = question.Element("ScoreCondition")?.Element("Verb")?.Element("Verb")?.Value;
                    string? ScoreConditionValue = question.Element("ScoreCondition")?.Element("Value")?.Value;
                    string? ScoreValue = question.Element("ScoreValue")?.Value;
                    string? UpperLimit = question.Element("UpperLimit")?.Value;
                    string? LowerLimit = question.Element("LowerLimit")?.Value;
                    string? UpperMaxLimit = question.Element("UpperMaxLimit")?.Value;
                    string? LowerMaxLimit = question.Element("LowerMaxLimit")?.Value;
                    string? ScoreMax = question.Element("ScoreMax")?.Value;
                    string? Name = question.Element("Name")?.Value;

                    // 将数据写入Excel
                    worksheet.Cells[row, 2].Value = procedureName;
                    worksheet.Cells[row, 3].Value = ScoreConditionExpression;
                    worksheet.Cells[row, 4].Value = ScoreConditionVerb;
                    worksheet.Cells[row, 5].Value = ScoreConditionValue;
                    worksheet.Cells[row, 6].Value = ScoreValue;
                    worksheet.Cells[row, 7].Value = UpperLimit;
                    worksheet.Cells[row, 8].Value = LowerLimit;
                    worksheet.Cells[row, 9].Value = UpperMaxLimit;
                    worksheet.Cells[row, 10].Value = LowerMaxLimit;
                    worksheet.Cells[row, 11].Value = ScoreMax;
                    worksheet.Cells[row, 12].Value = Name;

                    row++;
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
            string xlsxFile = Path.Combine(directory, $"{fileName}_Paper.xlsx");
            return xlsxFile;
        }
    }
}