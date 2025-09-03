using System;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace All_in_1
{
    internal class TaskOperationGenerator
    {
        public void Run()
        {
            Console.WriteLine("任务操作生成器");
            Console.WriteLine("请选择要上传的Excel文件：");
            string excelFilePath = Console.ReadLine();
            excelFilePath = excelFilePath.Replace("\"", "");

            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine("指定的Excel文件不存在");
                return;
            }

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                StringBuilder stateMachineOutput = new StringBuilder();

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var cells = worksheet.Cells[row, 1, row, 7];
                    string X = cells[1, 1].Text;
                    string Y = cells[1, 2].Text;
                    string A = cells[1, 3].Text;
                    string path2 = cells[1, 4].Text;
                    string path3 = cells[1, 5].Text;
                    string D = cells[1, 6].Text;
                    string valueValue = cells[1, 7].Text;

                    // 状态机部分
                    stateMachineOutput.AppendLine("<CustomStateMachine>");
                    stateMachineOutput.AppendLine("  <StateMachine>");
                    stateMachineOutput.AppendLine($"    <StateMachine name=\"C1-{X}\" sceneid=\"1\">");
                    stateMachineOutput.AppendLine("      <State name=\"Normal\" isDefaultState=\"\">");
                    stateMachineOutput.AppendLine("        <Action type=\"group\"></Action>");
                    stateMachineOutput.AppendLine("      </State>");
                    stateMachineOutput.AppendLine("      <State name=\"状态1\" preemptive=\"true\">");
                    stateMachineOutput.AppendLine("        <Action type=\"group\" seq=\"1\">");

                    if (A.Contains("hold")) stateMachineOutput.AppendLine("          <Action type=\"hold\" time=\"5\"/>");
                    if (A.Contains("active")) stateMachineOutput.AppendLine("          <Action type=\"active\" path=\"GMP_rjnj_texiao/GMP_ruanjiaonang_keli/GMP_rjnkl\" targetactive=\"1\" />");
                    if (A.Contains("anim")) stateMachineOutput.AppendLine("          <Action type=\"anim\" path=\"MZJC_zh/MZJC_JL_fm/JL_fm_jzf/jzf_sl16\" anim=\"kaiguan1\" startProgress=\"0\" endProgress=\"1\"/>");
                    if (A.Contains("rotate")) stateMachineOutput.AppendLine("          <Action type=\"rotate\" path=\"CJY_project/CJY_fm/CJY_fm/fm_zf/zf_299/zf_sl 124\" from=\"0\" to=\"-180\" speed=\"60\" axi=\"up\"/>");
                    if (A.Contains("video")) stateMachineOutput.AppendLine("          <Action type=\"video\" video=\"视频1.mp4\" op=\"play\" size=\"100,100\" waitforfinish=\"true\" showbar=\"true\"/>");
                    if (A.Contains("texture")) stateMachineOutput.AppendLine("          <Action type=\"texture\" texturename=\"头像1\" size=\"100,100\" bshow=\"true\" showClose=\"true\" uiinfo=\"uiinfo\" buttonText=\"按键1,按键2\"/>");
                    if (A.Contains("dlgshow")) stateMachineOutput.AppendLine("          <Action type=\"dlgshow\" name=\"dlgminimap\" show=\"1\" param=\"param\"/>");
                    if (A.Contains("tip")) stateMachineOutput.AppendLine("          <Action type=\"tip\" text=\"hello,welcome\" buttonText=\"按键1,按键2\" bshow=\"true\" bgSize=\"800,600\" useblock=\"true\" uiinfo=\"DuoXuanKuang\" alignmentmode=\"1\"/>");
                    if (A.Contains("move")) stateMachineOutput.AppendLine("          <Action type=\"move\" path=\"Main Camera\" pos=\"7.688791,1.119859,9.105009\" euler=\"0,0,0\" moveSpeed=\"1\" posOffset=\"0.1\" rotateSpeed=\"1\" rotOffset=\"1\" uniformMove=\"false\"/>");
                    if (A.Contains("processbar")) stateMachineOutput.AppendLine("           <Action type=\"processbar\" time=\"5\" tip=\"aaaa\"/>");

                    stateMachineOutput.AppendLine($"          <Action type=\"setouttervar\" name=\"DGJSC_M{Y}_STOP\" value=\"1\" />");
                    stateMachineOutput.AppendLine($"          <Action type=\"setouttervar\" name=\"DGJSC_M{(int.Parse(Y) + 1)}_INPUT\" value=\"1\" />");
                    stateMachineOutput.AppendLine("        </Action>");
                    stateMachineOutput.AppendLine("      </State>");
                    stateMachineOutput.AppendLine("      <Transition src=\"Normal\" dest=\"状态1\">");
                    stateMachineOutput.AppendLine("        <Condition type=\"and\">");
                    stateMachineOutput.AppendLine($"          <Condition type=\"oVar\" path=\"DGJSC_M{Y}_FLAG\" valvalue=\"1\" valtype=\"0\" valop=\"0\" />");
                    if (!string.IsNullOrEmpty(path2)) stateMachineOutput.AppendLine($"          <Condition type=\"objclick\" path=\"{path2}\" />");
                    if (!string.IsNullOrEmpty(path3)) stateMachineOutput.AppendLine($"          <Condition type=\"uiclick\" path=\"{path3}\" />");
                    if (!string.IsNullOrEmpty(D)) stateMachineOutput.AppendLine($"          <Condition type=\"roleposin\" destpos=\"{D}\" radius=\"1\" />");
                    stateMachineOutput.AppendLine("        </Condition>");
                    stateMachineOutput.AppendLine("      </Transition>");
                    stateMachineOutput.AppendLine("    </StateMachine>");
                    stateMachineOutput.AppendLine("  </StateMachine>");
                    stateMachineOutput.AppendLine("</CustomStateMachine>");

                    // 任务提示状态机部分
                    stateMachineOutput.AppendLine("<CustomStateMachine>");
                    stateMachineOutput.AppendLine("  <StateMachine>");
                    stateMachineOutput.AppendLine($"    <StateMachine name=\"S1-{X}\" sceneid=\"1\">");
                    stateMachineOutput.AppendLine("      <State name=\"showtip\" preemptive=\"true\">");
                    stateMachineOutput.AppendLine("        <Action type=\"group\" seq=\"0\">");
                    if (!string.IsNullOrEmpty(path2))
                    {
                        stateMachineOutput.AppendLine($"          <Action type=\"highlight\" path=\"{path2}\" color=\"0,255,0,255\" duration=\"-1\" constantOn=\"true\" />");
                    }
                    stateMachineOutput.AppendLine($"          <Action type=\"setlist\" listtype=\"string\" listname=\"DirectionList\" varname=\"direction\" value=\"{valueValue}\" op=\"1\" />");
                    stateMachineOutput.AppendLine("        </Action>");
                    stateMachineOutput.AppendLine("      </State>");
                    stateMachineOutput.AppendLine("      <State name=\"notshowtip\" isDefaultState=\"\"></State>");
                    stateMachineOutput.AppendLine("      <Transition src=\"showtip\" dest=\"notshowtip\">");
                    stateMachineOutput.AppendLine($"        <Condition type=\"oVar\" path=\"DGJSC_M{Y}_HINT\" valvalue=\"0\" valtype=\"0\" valop=\"0\" />");
                    stateMachineOutput.AppendLine("      </Transition>");
                    stateMachineOutput.AppendLine("      <Transition src=\"notshowtip\" dest=\"showtip\">");
                    stateMachineOutput.AppendLine("        <Condition type=\"and\">");
                    stateMachineOutput.AppendLine($"          <Condition type=\"oVar\" path=\"DGJSC_M{Y}_HINT\" valvalue=\"1\" valtype=\"0\" valop=\"0\" />");
                    stateMachineOutput.AppendLine("          <Condition type=\"operationmode\" modeid=\"0\" />");
                    stateMachineOutput.AppendLine("        </Condition>");
                    stateMachineOutput.AppendLine("      </Transition>");
                    stateMachineOutput.AppendLine("    </StateMachine>");
                    stateMachineOutput.AppendLine("  </StateMachine>");
                    stateMachineOutput.AppendLine("</CustomStateMachine>");
                }

                Console.WriteLine("生成的状态机：");
                Console.WriteLine(stateMachineOutput.ToString());
            }
        }
    }
}
