/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : OfficeTools.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using LKY_OfficeTools.Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;

namespace LKY_OfficeTools
{
    internal class OfficeTools
    {
        static void Main(string[] args)
        {
            //欢迎话术
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            
            //设置标题
            Console.Title = $"LKY Office Tools v{version}";

            //Header
            Console.WriteLine($"LKY Office Tools [版本 {version}]\n" +
                $"版权所有（C）LiuKaiyuan (Odysseus.Yuan)。保留所有权利。\n\n" +
                $"有关 {Console.Title} 更多信息，可发送邮件至：liukaiyuan@sjtu.edu.cn");

            /*
            Console.WriteLine($"******* Welcome to using LKY Office Tools v{version} *******");
            Console.WriteLine("******* Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc. *******");
            */

            //等待用户
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("\n请按 回车键 继续部署 ...");

            //权限检查
            Com_PrivilegeOS.PrivilegeAttention();

            if (Console.ReadKey().Key == ConsoleKey.Enter)
            {
                //继续
                new Lib_OfficeInstall();

                //退出机制
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write("\n请按任意键退出 ...");
                Console.ReadKey();
            }
            else
            {
                return;
            }
        }
    }
}
