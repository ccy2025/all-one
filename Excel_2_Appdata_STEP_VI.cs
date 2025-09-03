using System.Text;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Excel_2_Appdata_STEP_VI
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nExcel生成Appdata - HS-STEP_VI - StateMachine");
            Console.WriteLine("(StringBuilder)");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("仅支持标准HS-STEP_VI-AppData互转");
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

            // 创建一个与 Excel 文件名称相同的 XML 文件
            string xmlFilePath = Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(excelFilePath) + "_STEP_VI.xml");

            // 打开 Excel 文件
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // 读取第一个工作表
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // 使用StringBuilder构建XML
                // 手动方式构建，不够优雅，留着以后再改 2024.1.9
                StringBuilder xmlStringBuilder = new StringBuilder();

                string mod = worksheet.Cells[2, 1].Text;
                string name = worksheet.Cells[2, 2].Text;

                xmlStringBuilder.AppendLine("<AppData>\n");

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    string highLight = worksheet.Cells[row, 3].Text;
                    string setList = worksheet.Cells[row, 4].Text;
                    string stopoVar = worksheet.Cells[row, 5].Text;
                    string stopValvalue = worksheet.Cells[row, 6].Text;

                    xmlStringBuilder.AppendLine("\t<CustomStateMachine>");
                    xmlStringBuilder.AppendLine("\t\t<StateMachine>");
                    xmlStringBuilder.AppendLine($"\t\t\t<StateMachine name=\"{name}_STEP-{row - 1}\" sceneid=\"1\">\n");
                    xmlStringBuilder.AppendLine("\t\t\t\t<State name=\"normal\" isDefaultState=\"\">");
                    xmlStringBuilder.AppendLine("\t\t\t\t</State>\n");
                    xmlStringBuilder.AppendLine("\t\t\t\t<State name=\"show_inst\" preemptive=\"true\">");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t<Action type=\"group\" seq=\"0\">");
                    xmlStringBuilder.AppendLine($"\t\t\t\t\t\t<Action type=\"highlight\" path=\"{highLight}\" color=\"0,255,0,255\" duration=\"-1\" constantOn=\"true\" />");
                    xmlStringBuilder.AppendLine($"\t\t\t\t\t\t<Action type=\"setlist\" listtype=\"string\" listname=\"DirectionList\" varname=\"direction\" value=\"{setList}\" op=\"1\" />");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t</Action>");
                    xmlStringBuilder.AppendLine("\t\t\t\t</State>\n");
                    xmlStringBuilder.AppendLine("\t\t\t\t<State name=\"stop\" preemptive=\"true\">");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t<Action type=\"group\" seq=\"0\">");
                    xmlStringBuilder.AppendLine($"\t\t\t\t\t\t<Action type=\"setouttervar\" name=\"{mod}_{name}_STEP_VI\" value=\"{row - 1}\" />");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t</Action>");
                    xmlStringBuilder.AppendLine("\t\t\t\t</State>\n");
                    xmlStringBuilder.AppendLine("\t\t\t\t<Transition src=\"normal\" dest=\"show_inst\">");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t<Condition type=\"and\">");
                    xmlStringBuilder.AppendLine($"\t\t\t\t\t\t<Condition type=\"oVar\" path=\"{mod}_{name}_STEP_VI\" valvalue=\"{row - 2}\" valtype=\"0\" valop=\"0\" />");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t\t<Condition type=\"operationmode\" modeid=\"0\" />");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t</Condition>");
                    xmlStringBuilder.AppendLine("\t\t\t\t</Transition>\n");
                    xmlStringBuilder.AppendLine("\t\t\t\t<Transition src=\"show_inst\" dest=\"stop\">");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t<Condition type=\"and\">");
                    xmlStringBuilder.AppendLine($"\t\t\t\t\t\t<Condition type=\"oVar\" path=\"{mod}_{stopoVar}_VI\" valvalue=\"{stopValvalue}\" valtype=\"0\" valop=\"0\" />");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t</Condition>");
                    xmlStringBuilder.AppendLine("\t\t\t\t</Transition>\n");
                    xmlStringBuilder.AppendLine("\t\t\t\t<Transition src=\"stop\" dest=\"normal\">");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t<Condition type=\"and\">");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t\t<Condition type=\"auto\" />");
                    xmlStringBuilder.AppendLine("\t\t\t\t\t</Condition>");
                    xmlStringBuilder.AppendLine("\t\t\t\t</Transition>\n");
                    xmlStringBuilder.AppendLine("\t\t\t</StateMachine>");
                    xmlStringBuilder.AppendLine("\t\t</StateMachine>");
                    xmlStringBuilder.AppendLine("\t</CustomStateMachine>\n");

                }

                xmlStringBuilder.AppendLine("</AppData>");

                // 将构建好的XML字符串写入到文件
                File.WriteAllText(xmlFilePath, xmlStringBuilder.ToString());

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("XML文件已生成：" + xmlFilePath + "\n\n");
            }

            //Console.WriteLine("按任意键退出...");
            //Console.ReadKey();

        }

    }
}