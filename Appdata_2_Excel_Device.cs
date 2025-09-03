using System.Xml.Linq;
using OfficeOpenXml;

namespace All_in_1
{
    internal class Appdata_2_Excel_Device
    {
        public void Run()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nAppdata生成Excel - Device 10in1");
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

                // 写入首行表头
                worksheet.Cells["E1"].Value = "Device";
                worksheet.Cells["G1"].Value = "Button";
                worksheet.Cells["I1"].Value = "Switch";
                worksheet.Cells["M1"].Value = "Valve";
                worksheet.Cells["Q1"].Value = "SetValueButton";
                worksheet.Cells["R1"].Value = "LED";
                worksheet.Cells["T1"].Value = "SwitchLever";
                worksheet.Cells["W1"].Value = "Gage";
                worksheet.Cells["Y1"].Value = "Compass";
                worksheet.Cells["AE1"].Value = "OnOffLight";
                worksheet.Cells["AL1"].Value = "other";

                // 首行表头合并
                worksheet.Cells["A1:D1"].Merge = true;
                worksheet.Cells["E1:F1"].Merge = true;
                worksheet.Cells["G1:H1"].Merge = true;
                worksheet.Cells["I1:L1"].Merge = true;
                worksheet.Cells["M1:P1"].Merge = true;
                worksheet.Cells["R1:S1"].Merge = true;
                worksheet.Cells["T1:V1"].Merge = true;
                worksheet.Cells["W1:X1"].Merge = true;
                worksheet.Cells["Y1:AD1"].Merge = true;
                worksheet.Cells["AE1:AK1"].Merge = true;
                worksheet.Cells["AL1:AX1"].Merge = true;

                // 写入第二行表头
                worksheet.Cells[2, 1].Value = "ID";
                worksheet.Cells[2, 2].Value = "Path";
                worksheet.Cells[2, 3].Value = "Tip";
                worksheet.Cells[2, 4].Value = "Name";
                worksheet.Cells[2, 5].Value = "Type";
                worksheet.Cells[2, 6].Value = "VarName";
                worksheet.Cells[2, 7].Value = "OnColor";
                worksheet.Cells[2, 8].Value = "OffColor";
                worksheet.Cells[2, 9].Value = "SwitchAxi";
                worksheet.Cells[2, 10].Value = "SwitchSpeed";
                worksheet.Cells[2, 11].Value = "SwitchCloseDegree";
                worksheet.Cells[2, 12].Value = "SwitchOpenDegree";
                worksheet.Cells[2, 13].Value = "ValveMaxValue";
                worksheet.Cells[2, 14].Value = "ValveMinValueDegree";
                worksheet.Cells[2, 15].Value = "ValveDegreeRange";
                worksheet.Cells[2, 16].Value = "ValveAxi";
                worksheet.Cells[2, 17].Value = "SetButtonValue";
                worksheet.Cells[2, 18].Value = "Digit";
                worksheet.Cells[2, 19].Value = "Decimal";
                worksheet.Cells[2, 20].Value = "LeverStartOpenPos";
                worksheet.Cells[2, 21].Value = "LeverStartClosePos";
                worksheet.Cells[2, 22].Value = "LeverPathExchangeRule";
                worksheet.Cells[2, 23].Value = "GageMaxValue";
                worksheet.Cells[2, 24].Value = "GageColor";
                worksheet.Cells[2, 25].Value = "CompassZeroDegree";
                worksheet.Cells[2, 26].Value = "CompassMinValueDegree";
                worksheet.Cells[2, 27].Value = "CompassMaxDegreeRange";
                worksheet.Cells[2, 28].Value = "VarRange";
                worksheet.Cells[2, 29].Value = "CompassCoeff";
                worksheet.Cells[2, 30].Value = "CompassAxi";
                worksheet.Cells[2, 31].Value = "OnValue";
                worksheet.Cells[2, 32].Value = "OffValue";
                worksheet.Cells[2, 33].Value = "LightColor";
                worksheet.Cells[2, 34].Value = "LightDistance";
                worksheet.Cells[2, 35].Value = "Range";
                worksheet.Cells[2, 36].Value = "SpotAngle";
                worksheet.Cells[2, 37].Value = "Intensity";
                worksheet.Cells[2, 38].Value = "StateMachine";
                worksheet.Cells[2, 39].Value = "Key";
                worksheet.Cells[2, 40].Value = "TipFontSize";
                worksheet.Cells[2, 41].Value = "ColliderSize";
                worksheet.Cells[2, 42].Value = "ColliderCenter";
                worksheet.Cells[2, 43].Value = "TriggerDistance";
                worksheet.Cells[2, 44].Value = "Identifier";
                worksheet.Cells[2, 45].Value = "TransmitPos";
                worksheet.Cells[2, 46].Value = "TransmitLookAtPos";
                worksheet.Cells[2, 47].Value = "ObserveAxi";
                worksheet.Cells[2, 48].Value = "ObserveDistance";
                worksheet.Cells[2, 49].Value = "ObservePos";
                worksheet.Cells[2, 50].Value = "DeviceMapImage";

                // 行号
                int row = 3;

                // XML解析
                XDocument doc = XDocument.Load(xmlFile);

                // 不区分大小写遍历Device元素，避免不规范Device
                foreach (var device in doc.Descendants().Where(x => string.Equals(x.Name.LocalName, "Device", StringComparison.OrdinalIgnoreCase)))
                {
                    // Device通用项
                    string? ID = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ID", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Path = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Path", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Tip = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Tip", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Name = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))?.Value;

                    // Type通用项
                    string? Type = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Type", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? VarName = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "VarName", StringComparison.OrdinalIgnoreCase))?.Value;

                    // Button
                    string? OnColor = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "OnColor", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? OffColor = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "OffColor", StringComparison.OrdinalIgnoreCase))?.Value;

                    // Switch
                    string? SwitchAxi = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "SwitchAxi", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? SwitchSpeed = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "SwitchSpeed", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? SwitchCloseDegree = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "SwitchCloseDegree", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? SwitchOpenDegree = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "SwitchOpenDegree", StringComparison.OrdinalIgnoreCase))?.Value;

                    // Value
                    string? ValveMaxValue = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ValveMaxValue", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ValveMinValueDegree = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ValveMinValueDegree", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ValveDegreeRange = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ValveDegreeRange", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ValveAxi = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ValveAxi", StringComparison.OrdinalIgnoreCase))?.Value;

                    // SetValueButton
                    string? SetButtonValue = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "SetButtonValue", StringComparison.OrdinalIgnoreCase))?.Value;

                    // LED
                    string? Digit = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Digit", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Decimal = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Decimal", StringComparison.OrdinalIgnoreCase))?.Value;

                    // SwitchLever
                    string? LeverStartOpenPos = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "LeverStartOpenPos", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? LeverStartClosePos = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "LeverStartClosePos", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? LeverPathExchangeRule = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "LeverPathExchangeRule", StringComparison.OrdinalIgnoreCase))?.Value;

                    // Gage
                    string? GageMaxValue = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "GageMaxValue", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? GageColor = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "GageColor", StringComparison.OrdinalIgnoreCase))?.Value;

                    // Compass
                    string? CompassZeroDegree = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "CompassZeroDegree", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? CompassMinValueDegree = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "CompassMinValueDegree", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? CompassMaxDegreeRange = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "CompassMaxDegreeRange", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? VarRange = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "VarRange", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? CompassCoeff = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "CompassCoeff", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? CompassAxi = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "CompassAxi", StringComparison.OrdinalIgnoreCase))?.Value;

                    // OnOffLight
                    string? OnValue = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "OnValue", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? OffValue = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "OffValue", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? LightColor = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "LightColor", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? LightDistance = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "LightDistance", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Range = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Range", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? SpotAngle = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "SpotAngle", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Intensity = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Intensity", StringComparison.OrdinalIgnoreCase))?.Value;

                    // Other
                    string? StateMachine = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "StateMachine", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Key = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Key", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? TipFontSize = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "TipFontSize", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ColliderSize = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ColliderSize", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ColliderCenter = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ColliderCenter", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? TriggerDistance = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "TriggerDistance", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? Identifier = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Identifier", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? TransmitPos = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "TransmitPos", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? TransmitLookAtPos = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "TransmitLookAtPos", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ObserveAxi = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ObserveAxi", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ObserveDistance = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ObserveDistance", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? ObservePos = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "ObservePos", StringComparison.OrdinalIgnoreCase))?.Value;
                    string? DeviceMapImage = device.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "DeviceMapImage", StringComparison.OrdinalIgnoreCase))?.Value;

                    // 如果ID为空，跳过当前Device
                    //if (string.IsNullOrWhiteSpace(ID))
                    //{
                    //    continue;
                    //}

                    // 数据写入
                    worksheet.Cells[row, 1].Value = ID;
                    worksheet.Cells[row, 2].Value = Path;
                    worksheet.Cells[row, 3].Value = Tip;
                    worksheet.Cells[row, 4].Value = Name;
                    worksheet.Cells[row, 5].Value = Type;
                    worksheet.Cells[row, 6].Value = VarName;
                    worksheet.Cells[row, 7].Value = OnColor;
                    worksheet.Cells[row, 8].Value = OffColor;
                    worksheet.Cells[row, 9].Value = SwitchAxi;
                    worksheet.Cells[row, 10].Value = SwitchSpeed;
                    worksheet.Cells[row, 11].Value = SwitchCloseDegree;
                    worksheet.Cells[row, 12].Value = SwitchOpenDegree;
                    worksheet.Cells[row, 13].Value = ValveMaxValue;
                    worksheet.Cells[row, 14].Value = ValveMinValueDegree;
                    worksheet.Cells[row, 15].Value = ValveDegreeRange;
                    worksheet.Cells[row, 16].Value = ValveAxi;
                    worksheet.Cells[row, 17].Value = SetButtonValue;
                    worksheet.Cells[row, 18].Value = Digit;
                    worksheet.Cells[row, 19].Value = Decimal;
                    worksheet.Cells[row, 20].Value = LeverStartOpenPos;
                    worksheet.Cells[row, 21].Value = LeverStartClosePos;
                    worksheet.Cells[row, 22].Value = LeverPathExchangeRule;
                    worksheet.Cells[row, 23].Value = GageMaxValue;
                    worksheet.Cells[row, 24].Value = GageColor;
                    worksheet.Cells[row, 25].Value = CompassZeroDegree;
                    worksheet.Cells[row, 26].Value = CompassMinValueDegree;
                    worksheet.Cells[row, 27].Value = CompassMaxDegreeRange;
                    worksheet.Cells[row, 28].Value = VarRange;
                    worksheet.Cells[row, 29].Value = CompassCoeff;
                    worksheet.Cells[row, 30].Value = CompassAxi;
                    worksheet.Cells[row, 31].Value = OnValue;
                    worksheet.Cells[row, 32].Value = OffValue;
                    worksheet.Cells[row, 33].Value = LightColor;
                    worksheet.Cells[row, 34].Value = LightDistance;
                    worksheet.Cells[row, 35].Value = Range;
                    worksheet.Cells[row, 36].Value = SpotAngle;
                    worksheet.Cells[row, 37].Value = Intensity;
                    worksheet.Cells[row, 38].Value = StateMachine;
                    worksheet.Cells[row, 39].Value = Key;
                    worksheet.Cells[row, 40].Value = TipFontSize;
                    worksheet.Cells[row, 41].Value = ColliderSize;
                    worksheet.Cells[row, 42].Value = ColliderCenter;
                    worksheet.Cells[row, 43].Value = TriggerDistance;
                    worksheet.Cells[row, 44].Value = Identifier;
                    worksheet.Cells[row, 45].Value = TransmitPos;
                    worksheet.Cells[row, 46].Value = TransmitLookAtPos;
                    worksheet.Cells[row, 47].Value = ObserveAxi;
                    worksheet.Cells[row, 48].Value = ObserveDistance;
                    worksheet.Cells[row, 49].Value = ObservePos;
                    worksheet.Cells[row, 50].Value = DeviceMapImage;

                    row++;
                }

                // 自动调整列宽
                worksheet.Cells.AutoFitColumns();
                // 冻结前两行
                worksheet.View.FreezePanes(3, 1);
                // 保存xlsx文件
                workbook.SaveAs(new FileInfo(xlsxFile));
            }
        }

        static string GetOutputFilePath(string xmlFile)
        {
            string directory = Path.GetDirectoryName(xmlFile);
            string fileName = Path.GetFileNameWithoutExtension(xmlFile);
            string xlsxFile = Path.Combine(directory, $"{fileName}_Device_10in1.xlsx");
            return xlsxFile;
        }
    }
}