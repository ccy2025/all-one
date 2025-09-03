using All_in_1;

class Progress
{
    static void Main(string[] args)
    {
        Console.BackgroundColor = ConsoleColor.DarkGreen;
        Console.WriteLine(".NET8_14in1_2024.4.29\n");
        Console.ResetColor();

        while (true)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("请选择要运行的程序：");
            Console.ResetColor();

            Console.WriteLine("[1] - Appdata生成Excel - StateMachine *TEST");
            Console.WriteLine("[2] - Appdata生成Excel - setlist");
            Console.WriteLine("[3] - Paper生成Excel");
            Console.WriteLine("[4] - Excel生成Paper");
            Console.WriteLine("[5] - Appdata生成Excel - Device");
            Console.WriteLine("[6] - Appdata生成Excel - Device 10in1");
            Console.WriteLine("[7] - Excel生成Appdata - Device");
            Console.WriteLine("[8] - Excel生成Appdata - HS-STEP_VI - StateMachine");
            Console.WriteLine("[9] - Appdata生成Excel - HS-STEP_VI - StateMachine");
            Console.WriteLine("[10*] - Excel生成Appdata - HS-STEP_VI - Add EXAM-StateMachine");
            //Console.WriteLine("11 - Appdata生成Excel - State");
            //Console.WriteLine("12 - Appdata生成Excel - Transition");
            Console.WriteLine("[13] - Task生成Excel - content & tip & paramkey");
            Console.WriteLine("[16] - Excel生成Appdata - TaskOperationGenerator"); // 新增选项
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("任意其他键 - 退出");
            Console.ResetColor();

            Console.Write("请输入选项：");
            Console.ForegroundColor = ConsoleColor.Cyan;

            var choice = Console.ReadLine();

            Console.Clear();

            switch (choice)
            {
                case "1":
                    var program_Appdata_2_Excel = new Appdata_2_Excel_StateMachine();
                    program_Appdata_2_Excel.Run();
                    break;

                case "2":
                    var program_Appdata_2_Excel_setlist = new Appdata_2_Excel_setlist();
                    program_Appdata_2_Excel_setlist.Run();
                    break;

                case "3":
                    var program_Paper_2_Excel = new Paper_2_Excel();
                    program_Paper_2_Excel.Run();
                    break;

                case "4":
                    var program_Excel_2_Paper = new Excel_2_Paper();
                    program_Excel_2_Paper.Run();
                    break;

                case "5":
                    var program_Appdata_2_Excel_Device_Lite = new Appdata_2_Excel_Device_Lite();
                    program_Appdata_2_Excel_Device_Lite.Run();
                    break;

                case "6":
                    var program_Appdata_2_Excel_Device = new Appdata_2_Excel_Device();
                    program_Appdata_2_Excel_Device.Run();
                    break;

                case "7":
                    var program_Excel_2_Appdata_Device_2 = new Excel_2_Appdata_Device_2();
                    program_Excel_2_Appdata_Device_2.Run();
                    break;

                case "7.5":
                    var program_Excel_2_Appdata_Device = new Excel_2_Appdata_Device();
                    program_Excel_2_Appdata_Device.Run();
                    break;

                case "8":
                    var program_Excel_2_Appdata_STEP_VI = new Excel_2_Appdata_STEP_VI();
                    program_Excel_2_Appdata_STEP_VI.Run();
                    break;

                case "9":
                    var program_Appdata_2_Excel_STEP_VI = new Appdata_2_Excel_STEP_VI();
                    program_Appdata_2_Excel_STEP_VI.Run();
                    break;

                //case "11":
                //    var program_Appdata_2_Excel_State = new Appdata_2_Excel_State();
                //    program_Appdata_2_Excel_State.Run();
                //    break;

                //case "12":
                //    var program_Appdata_2_Excel_Transition = new Appdata_2_Excel_Transition();
                //    program_Appdata_2_Excel_Transition.Run();
                //    break;

                case "13":
                    var program_Task_2_Excel_CTP = new Task_2_Excel_CTP();
                    program_Task_2_Excel_CTP.Run();
                    break;

                case "16": // 新增功能的case
                    var program_TaskOperationGenerator = new TaskOperationGenerator();
                    program_TaskOperationGenerator.Run();
                    break;

                default:
                    return;

            }
        }
    }

}