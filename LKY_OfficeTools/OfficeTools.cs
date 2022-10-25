/*
 *      [LKY Common Tools] Copyright (C) 2022 SJTU Inc.
 *      
 *      FileName : OfficeTools.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace LKY_OfficeTools
{
    internal class OfficeTools
    {
        static void Main(string[] args)
        {
            //欢迎话术
            Console.Write("******* 欢迎使用 LKY Office Tools，请按 回车键 继续 ...");

            if (Console.ReadKey().Key == ConsoleKey.Enter)
            {
                //继续
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.Write("\n\n------> 正在获取最新版 Microsoft Office 版本 ...\n");
                new Lib_OfficeInfo();

                Console.ReadKey();
            }
            else
            {
                return;
            }
        }
    }
}
