/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 - 2023 OdysseusYuan@foxmail.com Inc.
 *      
 *      FileName : Lib_AppMessage.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.Threading;
using static LKY_OfficeTools.Common.Com_Timer;
using static LKY_OfficeTools.Lib.Lib_AppCommand;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_AppMessage
    {
        internal class KeyMsg
        {
            internal static void Quit(int exit_code)
            {
                //清理SDK缓存
                Lib_AppSdk.Clean();

                //退出机制
                if (!AppCommandFlag.HasFlag(ArgsFlag.None_Finish_PressKey))
                {
                    //不包含“结束无需确认”命令行，需要人工按键结束
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.Write("\n请按 任意键 退出 ...");
                    Console.ReadKey();

                    Environment.Exit(exit_code);
                }
            }

            internal static bool Confirm(string msg_str = null)
            {
                Console.ForegroundColor = ConsoleColor.Gray;

                //判断是否为空
                if (string.IsNullOrWhiteSpace(msg_str))
                {
                    //msg为空，直接展示回车键继续，并且前面不空格
                    msg_str = $"\n请按 回车键（Enter）继续 ...";
                }
                else
                {
                    //msg不为空，一般在运行过程中的确认，有空格，并且增加逗号
                    msg_str = $"        {msg_str}，请按 回车键（Enter）继续 ...";
                }

                Console.Write(msg_str);     //提示信息
                new Log(msg_str, ConsoleColor.Gray, Log.Output_Type.Write);     //写入日志

                if (Console.ReadKey().Key == ConsoleKey.Enter)
                {
                    Console.WriteLine();    //增加一个空白行
                    return true;
                }
                else
                {
                    Console.WriteLine();    //增加一个空白行
                    return false;
                }
            }

            internal static void DoByTime(string msg_str, int countdown_time)
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine();    //插入一个空白行

                //设置一个倒计时组件
                Countdown_Timer timer = new Countdown_Timer();
                timer.Start(countdown_time);

                //循环输出
                while (timer.isRun)
                {
                    string msg = $"{msg_str}将在 {timer.Remaining_Time} 秒内开始 ...";
                    new Log(msg, ConsoleColor.Gray, Log.Output_Type.Write);     //写入日志
                    Thread.Sleep(100);

                    //输出消息倒计时
                    Console.Write($"\r{msg}");
                }

                //倒计时结束后，告知开始
                Console.Write($"\r{msg_str}启动 ...                ");

                //完成等待
                Console.WriteLine();    //增加一个空白行
                return;
            }

            internal static bool Choose(string todo_thing)
            {
                new Log($"\n     ★ {todo_thing}", ConsoleColor.Gray);

                Console.ForegroundColor = ConsoleColor.Gray;
                string msg = $"        按 回车键（Enter）确认执行上述操作，按 其它键 跳过此环节 ...";

                Console.Write(msg);
                new Log(msg, ConsoleColor.Gray, Log.Output_Type.Write);     //写入日志

                if (Console.ReadKey().Key == ConsoleKey.Enter)
                {
                    Console.WriteLine();    //增加一个空白行
                    return true;
                }
                else
                {
                    Console.WriteLine();    //增加一个空白行
                    return false;
                }
            }
        }
    }
}
