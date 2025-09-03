using System.Xml;
using OfficeOpenXml;

namespace All_in_1
{
    internal class bkp_temp
    {
        //static void Main(string[] args)
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nExcel生成Paper");
            Console.ResetColor();

            Console.WriteLine("请输入Excel文件路径：");
            string excelFilePath = Console.ReadLine();

            //去除路径空格自动双引号
            //string newexcelFilePath = null;
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
            string xmlFilePath = Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(excelFilePath) + ".xml");

            // 打开 Excel 文件
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // 创建 XML 文档
                XmlDocument xmlDoc = new XmlDocument();
                XmlElement rootElement = xmlDoc.CreateElement("AppData");
                xmlDoc.AppendChild(rootElement);


                string cellValue = worksheet.Cells[2, 1].Value?.ToString();
                XmlElement ProjectnameElement = xmlDoc.CreateElement("ProjectName");
                ProjectnameElement.InnerText = cellValue;
                rootElement.AppendChild(ProjectnameElement);


                int rowCount = worksheet.Dimension.Rows;
                int procedureIdCounter = 1;
                int questionIdCounter = 1;


                // 使用字典来跟踪已经创建的Procedure元素
                Dictionary<string, XmlElement> procedureDict = new Dictionary<string, XmlElement>();
                XmlElement currentProcedureElement = null;

                // 遍历第一列、第二列和第三列数据，将其添加到 XML 中
                for (int i = 2; i <= rowCount; i++)
                {
                    string procedureName = worksheet.Cells[i, 2].Value?.ToString();
                    string questionName = worksheet.Cells[i, 3].Value?.ToString();
                    string questionDescription = worksheet.Cells[i, 4].Value?.ToString();
                    string questionValue = worksheet.Cells[i, 5].Value?.ToString();

                    if (!string.IsNullOrEmpty(procedureName))
                    {
                        if (!procedureDict.ContainsKey(procedureName))
                        {
                            // 如果Procedure元素不存在，则创建一个新的Procedure元素
                            XmlElement procedureElement = xmlDoc.CreateElement("Procedures");

                            //XmlElement idElement = xmlDoc.CreateElement("ID");
                            //idElement.InnerText = "P" + procedureIdCounter++;
                            //procedureElement.AppendChild(idElement);

                            XmlElement idElement = xmlDoc.CreateElement("ID");
                            string procedureId = "P" + procedureIdCounter++;
                            idElement.InnerText = procedureId;
                            procedureElement.AppendChild(idElement);

                            XmlElement procGroupIdElement = xmlDoc.CreateElement("ProcGroupID");
                            procGroupIdElement.InnerText = "-1";
                            procedureElement.AppendChild(procGroupIdElement);

                            XmlElement nameElement = xmlDoc.CreateElement("Name");
                            nameElement.InnerText = procedureName;
                            procedureElement.AppendChild(nameElement);

                            //Procedure通用内容
                            XmlElement startConditionElement = xmlDoc.CreateElement("StartCondition");
                            XmlElement startExpressionElement = xmlDoc.CreateElement("Expression");
                            startExpressionElement.InnerText = "default";
                            startConditionElement.AppendChild(startExpressionElement);
                            XmlElement startVerbElement = xmlDoc.CreateElement("Verb");
                            XmlElement startInnerVerbElement = xmlDoc.CreateElement("Verb");
                            startInnerVerbElement.InnerText = "OR";
                            startVerbElement.AppendChild(startInnerVerbElement);
                            startConditionElement.AppendChild(startVerbElement);
                            XmlElement startValueElement = xmlDoc.CreateElement("Value");
                            startValueElement.InnerText = "0";
                            startConditionElement.AppendChild(startValueElement);
                            XmlElement startDescElement = xmlDoc.CreateElement("Desc");
                            startDescElement.InnerText = "default description";
                            startConditionElement.AppendChild(startDescElement);
                            XmlElement startSatisfiedElement = xmlDoc.CreateElement("Satisfied");
                            startSatisfiedElement.InnerText = "False";
                            startConditionElement.AppendChild(startSatisfiedElement);
                            procedureElement.AppendChild(startConditionElement);

                            XmlElement endConditionElement = xmlDoc.CreateElement("EndCondition");
                            XmlElement endExpressionElement = xmlDoc.CreateElement("Expression");
                            endExpressionElement.InnerText = "default";
                            endConditionElement.AppendChild(endExpressionElement);
                            XmlElement endVerbElement = xmlDoc.CreateElement("Verb");
                            XmlElement endInnerVerbElement = xmlDoc.CreateElement("Verb");
                            endInnerVerbElement.InnerText = "OR";
                            endVerbElement.AppendChild(endInnerVerbElement);
                            endConditionElement.AppendChild(endVerbElement);
                            XmlElement endValueElement = xmlDoc.CreateElement("Value");
                            endValueElement.InnerText = "0";
                            endConditionElement.AppendChild(endValueElement);
                            XmlElement endDescElement = xmlDoc.CreateElement("Desc");
                            endDescElement.InnerText = "default description";
                            endConditionElement.AppendChild(endDescElement);
                            XmlElement endSatisfiedElement = xmlDoc.CreateElement("Satisfied");
                            endSatisfiedElement.InnerText = "False";
                            endConditionElement.AppendChild(endSatisfiedElement);
                            procedureElement.AppendChild(endConditionElement);

                            XmlElement isStartedElement = xmlDoc.CreateElement("IsStarted");
                            isStartedElement.InnerText = "False";
                            procedureElement.AppendChild(isStartedElement);
                            XmlElement isFinishedElement = xmlDoc.CreateElement("IsFinished");
                            isFinishedElement.InnerText = "False";
                            procedureElement.AppendChild(isFinishedElement);
                            XmlElement descElement = xmlDoc.CreateElement("Desc");
                            descElement.InnerText = "default description";
                            procedureElement.AppendChild(descElement);
                            XmlElement stateElement = xmlDoc.CreateElement("State");
                            stateElement.InnerText = "valid";
                            procedureElement.AppendChild(stateElement);

                            rootElement.AppendChild(procedureElement);

                            // 将新创建的Procedure元素添加到字典中
                            procedureDict[procedureName] = procedureElement;

                            currentProcedureElement = procedureElement;
                        }
                    }

                    if (!string.IsNullOrEmpty(questionName))
                    {
                        // 获取当前 Procedure 的 ProcedureID
                        string procedureId = procedureDict.FirstOrDefault(x => x.Key == procedureName).Value?.SelectSingleNode("ID")?.InnerText;

                        // 创建 Questions 元素
                        XmlElement questionsElement = xmlDoc.CreateElement("Questions");
                        rootElement.AppendChild(questionsElement);

                        XmlElement questionIdElement = xmlDoc.CreateElement("ID");
                        string questionId = "T" + questionIdCounter++;
                        questionIdElement.InnerText = questionId;
                        questionsElement.AppendChild(questionIdElement);

                        XmlElement questionNameElement = xmlDoc.CreateElement("Name");
                        questionNameElement.InnerText = questionName;
                        questionsElement.AppendChild(questionNameElement);

                        // 添加 <ProcedureID> 元素，其值为当前 Procedure 的 ProcedureID
                        XmlElement procedureidElement = xmlDoc.CreateElement("ProcedureID");
                        procedureidElement.InnerText = procedureId;
                        questionsElement.AppendChild(procedureidElement);

                        //XmlElement sceElement = xmlDoc.CreateElement("ScoreCondition");
                        //XmlElement sceNameElement = xmlDoc.CreateElement("Expression");
                        //sceNameElement.InnerText = questionDescription;
                        //sceElement.AppendChild(sceNameElement);
                        //questionsElement.AppendChild(sceElement);

                        //通用设置
                        //----------------------------
                        //---------------------------
                        // 添加<SectionID>元素
                        XmlElement sectionIdElement = xmlDoc.CreateElement("SectionID");
                        sectionIdElement.InnerText = "-1";
                        questionsElement.AppendChild(sectionIdElement);

                        // 添加<DeviceID>元素
                        XmlElement deviceIdElement = xmlDoc.CreateElement("DeviceID");
                        deviceIdElement.InnerText = "-1";
                        questionsElement.AppendChild(deviceIdElement);

                        // 添加<PlayerID>元素
                        XmlElement playerIdElement = xmlDoc.CreateElement("PlayerID");
                        playerIdElement.InnerText = "-1";
                        questionsElement.AppendChild(playerIdElement);

                        // 添加<ScoreCondition>元素
                        XmlElement scoreConditionElement = xmlDoc.CreateElement("ScoreCondition");
                        //----------------------------------------------------------------------------
                        XmlElement sceNameElement = xmlDoc.CreateElement("Expression");
                        sceNameElement.InnerText = questionDescription;
                        scoreConditionElement.AppendChild(sceNameElement);
                        questionsElement.AppendChild(scoreConditionElement);
                        //------------------------------------------------------------------------------
                        XmlElement scoreVerbElement = xmlDoc.CreateElement("Verb");
                        XmlElement scoreInnerVerbElement = xmlDoc.CreateElement("Verb");
                        scoreInnerVerbElement.InnerText = ">=";
                        scoreVerbElement.AppendChild(scoreInnerVerbElement);
                        scoreConditionElement.AppendChild(scoreVerbElement);

                        XmlElement scoreValueElement = xmlDoc.CreateElement("Value");
                        //scoreValueElement.InnerText = "0.5";
                        scoreValueElement.InnerText = questionValue;
                        scoreConditionElement.AppendChild(scoreValueElement);

                        XmlElement timeRelationShipElement = xmlDoc.CreateElement("TimeRelationShip");
                        XmlElement relateTypeElement = xmlDoc.CreateElement("RelateType");
                        relateTypeElement.InnerText = "NOTIME";
                        timeRelationShipElement.AppendChild(relateTypeElement);
                        XmlElement intervalElement = xmlDoc.CreateElement("Interval");
                        intervalElement.InnerText = "0";
                        timeRelationShipElement.AppendChild(intervalElement);
                        scoreConditionElement.AppendChild(timeRelationShipElement);

                        XmlElement scoreSatisfiedElement = xmlDoc.CreateElement("Satisfied");
                        scoreSatisfiedElement.InnerText = "False";
                        scoreConditionElement.AppendChild(scoreSatisfiedElement);

                        XmlElement orderRelateElement = xmlDoc.CreateElement("OrderRelate");
                        orderRelateElement.InnerText = "NO";
                        scoreConditionElement.AppendChild(orderRelateElement);

                        XmlElement scoreDescElement = xmlDoc.CreateElement("Desc");
                        scoreDescElement.InnerText = "default description";
                        scoreConditionElement.AppendChild(scoreDescElement);

                        // 2024新增
                        XmlElement scoreParamsElement = xmlDoc.CreateElement("Params");
                        scoreParamsElement.InnerText = questionDescription;
                        scoreConditionElement.AppendChild(scoreParamsElement);

                        questionsElement.AppendChild(scoreConditionElement);

                        // 添加<CurrentScoreValue>元素
                        XmlElement currentScoreValueElement = xmlDoc.CreateElement("CurrentScoreValue");
                        currentScoreValueElement.InnerText = "0";
                        questionsElement.AppendChild(currentScoreValueElement);

                        // 添加<ScoreValue>元素
                        XmlElement scoreValueElement2 = xmlDoc.CreateElement("ScoreValue");
                        scoreValueElement2.InnerText = "1";
                        questionsElement.AppendChild(scoreValueElement2);

                        // 添加<ScoreMax>元素
                        XmlElement scoreMaxElement = xmlDoc.CreateElement("ScoreMax");
                        scoreMaxElement.InnerText = "5";
                        questionsElement.AppendChild(scoreMaxElement);

                        // 添加<ScoreInterval>元素
                        XmlElement scoreIntervalElement = xmlDoc.CreateElement("ScoreInterval");
                        scoreIntervalElement.InnerText = "1";
                        questionsElement.AppendChild(scoreIntervalElement);

                        // 添加<TotalScoreTime>元素
                        XmlElement totalScoreTimeElement = xmlDoc.CreateElement("TotalScoreTime");
                        totalScoreTimeElement.InnerText = "0";
                        questionsElement.AppendChild(totalScoreTimeElement);

                        // 添加<StartCondition>元素
                        XmlElement startConditionElement = xmlDoc.CreateElement("StartCondition");

                        XmlElement startExpressionElement = xmlDoc.CreateElement("Expression");
                        startExpressionElement.InnerText = "default";
                        startConditionElement.AppendChild(startExpressionElement);

                        XmlElement startVerbElement = xmlDoc.CreateElement("Verb");
                        XmlElement startInnerVerbElement = xmlDoc.CreateElement("Verb");
                        startInnerVerbElement.InnerText = "OR";
                        startVerbElement.AppendChild(startInnerVerbElement);
                        startConditionElement.AppendChild(startVerbElement);

                        XmlElement startValueElement = xmlDoc.CreateElement("Value");
                        startValueElement.InnerText = "0";
                        startConditionElement.AppendChild(startValueElement);

                        XmlElement startDescElement = xmlDoc.CreateElement("Desc");
                        startDescElement.InnerText = "default description";
                        startConditionElement.AppendChild(startDescElement);

                        XmlElement startSatisfiedElement = xmlDoc.CreateElement("Satisfied");
                        startSatisfiedElement.InnerText = "False";
                        startConditionElement.AppendChild(startSatisfiedElement);

                        questionsElement.AppendChild(startConditionElement);

                        // 添加<EndCondition>元素
                        XmlElement endConditionElement = xmlDoc.CreateElement("EndCondition");

                        XmlElement endExpressionElement = xmlDoc.CreateElement("Expression");
                        endExpressionElement.InnerText = "default";
                        endConditionElement.AppendChild(endExpressionElement);

                        XmlElement endVerbElement = xmlDoc.CreateElement("Verb");
                        XmlElement endInnerVerbElement = xmlDoc.CreateElement("Verb");
                        endInnerVerbElement.InnerText = "OR";
                        endVerbElement.AppendChild(endInnerVerbElement);
                        endConditionElement.AppendChild(endVerbElement);

                        XmlElement endValueElement = xmlDoc.CreateElement("Value");
                        endValueElement.InnerText = "0";
                        endConditionElement.AppendChild(endValueElement);

                        XmlElement endDescElement = xmlDoc.CreateElement("Desc");
                        endDescElement.InnerText = "default description";
                        endConditionElement.AppendChild(endDescElement);

                        XmlElement endSatisfiedElement = xmlDoc.CreateElement("Satisfied");
                        endSatisfiedElement.InnerText = "False";
                        endConditionElement.AppendChild(endSatisfiedElement);

                        questionsElement.AppendChild(endConditionElement);

                        // 添加<UpperLimit>元素
                        XmlElement upperLimitElement = xmlDoc.CreateElement("UpperLimit");
                        upperLimitElement.InnerText = "0";
                        questionsElement.AppendChild(upperLimitElement);

                        // 添加<UpperMaxLimit>元素
                        XmlElement upperMaxLimitElement = xmlDoc.CreateElement("UpperMaxLimit");
                        upperMaxLimitElement.InnerText = "0";
                        questionsElement.AppendChild(upperMaxLimitElement);

                        // 添加<LowerLimit>元素
                        XmlElement lowerLimitElement = xmlDoc.CreateElement("LowerLimit");
                        lowerLimitElement.InnerText = "0";
                        questionsElement.AppendChild(lowerLimitElement);

                        // 添加<LowerMaxLimit>元素
                        XmlElement lowerMaxLimitElement = xmlDoc.CreateElement("LowerMaxLimit");
                        lowerMaxLimitElement.InnerText = "0";
                        questionsElement.AppendChild(lowerMaxLimitElement);

                        // 添加<OrderRelateQuestionID>元素
                        XmlElement orderRelateQuestionIdElement = xmlDoc.CreateElement("OrderRelateQuestionID");
                        orderRelateQuestionIdElement.InnerText = "-1";
                        questionsElement.AppendChild(orderRelateQuestionIdElement);

                        // 添加<TimeRelateQuestionID>元素
                        XmlElement timeRelateQuestionIdElement = xmlDoc.CreateElement("TimeRelateQuestionID");
                        timeRelateQuestionIdElement.InnerText = "-1";
                        questionsElement.AppendChild(timeRelateQuestionIdElement);

                        // 添加<StartTime>元素
                        XmlElement startTimeElement = xmlDoc.CreateElement("StartTime");
                        startTimeElement.InnerText = "0";
                        questionsElement.AppendChild(startTimeElement);

                        // 添加<StartDuration>元素
                        XmlElement startDurationElement = xmlDoc.CreateElement("StartDuration");
                        startDurationElement.InnerText = "0";
                        questionsElement.AppendChild(startDurationElement);

                        // 添加<FinishTime>元素
                        XmlElement finishTimeElement = xmlDoc.CreateElement("FinishTime");
                        finishTimeElement.InnerText = "0";
                        questionsElement.AppendChild(finishTimeElement);

                        // 添加<FinishDuration>元素
                        XmlElement finishDurationElement = xmlDoc.CreateElement("FinishDuration");
                        finishDurationElement.InnerText = "0";
                        questionsElement.AppendChild(finishDurationElement);

                        // 添加<StartConditionConcerned>元素
                        XmlElement startConditionConcernedElement = xmlDoc.CreateElement("StartConditionConcerned");
                        startConditionConcernedElement.InnerText = "False";
                        questionsElement.AppendChild(startConditionConcernedElement);

                        // 添加<IsStarted>元素
                        XmlElement isStartedElement = xmlDoc.CreateElement("IsStarted");
                        isStartedElement.InnerText = "False";
                        questionsElement.AppendChild(isStartedElement);

                        // 添加<IsFinished>元素
                        XmlElement isFinishedElement = xmlDoc.CreateElement("IsFinished");
                        isFinishedElement.InnerText = "False";
                        questionsElement.AppendChild(isFinishedElement);

                        XmlElement descElement = xmlDoc.CreateElement("Desc");
                        descElement.InnerText = questionName;
                        questionsElement.AppendChild(descElement);

                        // 添加<State>元素
                        XmlElement stateElement = xmlDoc.CreateElement("State");
                        stateElement.InnerText = "valid";
                        questionsElement.AppendChild(stateElement);

                        // 添加<HasScored>元素
                        XmlElement hasScoredElement = xmlDoc.CreateElement("HasScored");
                        hasScoredElement.InnerText = "False";
                        questionsElement.AppendChild(hasScoredElement);

                        //-------------------------------------
                        // 2024新增
                        // 添加<IsStartCond>元素
                        XmlElement isStartCondElement = xmlDoc.CreateElement("IsStartCond");
                        isStartCondElement.InnerText = "False";
                        questionsElement.AppendChild(isStartCondElement);

                        // 添加<IsRedoEvaluate>元素
                        XmlElement isRedoEvaluateElement = xmlDoc.CreateElement("IsRedoEvaluate");
                        isRedoEvaluateElement.InnerText = "False";
                        questionsElement.AppendChild(isRedoEvaluateElement);



                        // 添加 QuestionID 元素到对应的 Procedure 节点
                        if (currentProcedureElement != null)
                        {
                            XmlElement questionIdParentElement = xmlDoc.CreateElement("Questions");
                            currentProcedureElement.AppendChild(questionIdParentElement);

                            XmlElement questionIdChildElement = xmlDoc.CreateElement("ID");
                            questionIdChildElement.InnerText = questionId;
                            questionIdParentElement.AppendChild(questionIdChildElement);

                        }
                    }
                }

                XmlElement LocaleElement = xmlDoc.CreateElement("Locale");
                LocaleElement.InnerText = "简体中文";
                rootElement.AppendChild(LocaleElement);

                // 2024新增同变量原则
                XmlElement IsUseRulesElement = xmlDoc.CreateElement("IsUseRules");
                IsUseRulesElement.InnerText = "True";
                rootElement.AppendChild(IsUseRulesElement);

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