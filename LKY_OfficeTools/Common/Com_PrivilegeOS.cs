/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_PrivilegeOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Security.Principal;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 权限相关类库
    /// </summary>
    internal class Com_PrivilegeOS
    {
        /// <summary>
        /// 判断自身程序是否以管理员身份运行
        /// </summary>
        /// <returns></returns>
        internal static bool IsRunByAdmin()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        /// <summary>
        /// 权限处理结果
        /// </summary>
        internal static void PrivilegeAttention()
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine($"\n------> 正在进行 权限检查 ...");

            //提权检验，非提权，激活会出问题
            if (!IsRunByAdmin())
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"     × 权限错误，请以管理员身份运行此文件！");
                Console.Read();
                Environment.Exit(-1);
            }

            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine($"     √ 权限检查通过。");
        }
    }
}
