/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_PrivilegeOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Security.Principal;
using static LKY_OfficeTools.Lib.Lib_AppState;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppReport;
using LKY_OfficeTools.Lib;
using static LKY_OfficeTools.Lib.Lib_AppInfo;

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
            new Log($"\n------> 正在进行 {AppAttribute.AppName} 权限检查 ...", ConsoleColor.DarkCyan);

            //提权检验，非提权，激活会出问题
            if (!IsRunByAdmin())
            {
                new Log($"     × 权限错误，请以管理员身份运行此文件！", ConsoleColor.DarkRed);

                Current_StageType = ProcessStage.Finish_Fail;     //设置为失败模式
                Pointing(ProcessStage.Finish_Fail);  //回收

                //退出提示
                KeyMsg.Quit();

                Environment.Exit(-1);
            }

            new Log($"     √ 已通过 {AppAttribute.AppName} 权限检查。", ConsoleColor.DarkGreen);
        }
    }
}
