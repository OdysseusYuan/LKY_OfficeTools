/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppMessage.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using static LKY_OfficeTools.Lib.Lib_AppCommand;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 用户消息提示类库
    /// </summary>
    internal class Lib_AppMessage
    {
        /// <summary>
        /// 生成一条与按键相关的信息的类库
        /// </summary>
        internal class KeyMsg
        {
            /// <summary>
            /// 按键退出 & 完成善后
            /// </summary>
            internal static void Quit()
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
                }
            }

            /// <summary>
            /// 需要用户确认的信息（按 回车键 继续）
            /// </summary>
            internal static bool Confirm(string msg_str = null)
            {
                Console.ForegroundColor = ConsoleColor.Gray;

                //判断是否为空
                if (string.IsNullOrWhiteSpace(msg_str))
                {
                    //msg为空，直接展示回车键继续，并且前面不空格
                    Console.Write($"\n请按 回车键 继续 ...");
                }
                else
                {
                    //msg不为空，一般在运行过程中的确认，有空格，并且增加逗号
                    Console.Write($"        {msg_str}，请按 回车键 继续 ...");
                }

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
