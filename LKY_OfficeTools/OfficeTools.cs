/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : OfficeTools.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
            Version version = Assembly.GetExecutingAssembly().GetName().Version;

            Console.WriteLine($"******* Welcome to using LKY Office Tools v{version} *******");
            Console.WriteLine("******* Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc. *******");

            Console.Write("\n请按 回车键 继续 ...");

            if (Console.ReadKey().Key == ConsoleKey.Enter)
            {
                //继续
                new Lib_OfficeInstall();

                Console.ReadKey();
            }
            else
            {
                return;
            }
        }
    }
}
