/*
 *      [LKY Office Tools] Copyright (C) 2022 - 2024 LiuKaiyuan Inc.
 *      
 *      FileName : Com_PrivilegeOS.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.IO;
using System.Security.Principal;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools.Common
{
    internal class Com_PrivilegeOS
    {
        internal static bool IsRunByAdmin()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        internal static bool CanWriteProgramDataDir()
        {
            try
            {
                string time_sn = $"{DateTime.UtcNow.Hour}{DateTime.UtcNow.Minute}{DateTime.UtcNow.Second}{DateTime.UtcNow.Millisecond}";
                string test_file_dir = AppPath.Documents.Temp + $"\\{time_sn}";
                string test_file_path = test_file_dir + "\\time_sn.tmp";
                var isSuccess = Com_FileOS.Write.TextToFile(test_file_path, "test", false);

                //删除测试的文件和目录
                if (Directory.Exists(test_file_dir))
                {
                    Directory.Delete(test_file_dir, true);
                }

                return isSuccess;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }

        internal static void PrivilegeAttention()
        {
            new Log($"\n------> 正在进行 {AppAttribute.AppName} 权限检查 ...", ConsoleColor.DarkCyan);

            //提权检验，非提权，激活会出问题
            if (!IsRunByAdmin())
            {
                new Log($"     × 权限错误，请以管理员身份运行此文件！", ConsoleColor.DarkRed);

                Current_StageType = ProcessStage.Finish_Fail;     //设置为失败模式

                //退出提示
                KeyMsg.Quit(-1);
            }

            //ProgramData目录权限检查
            if (!CanWriteProgramDataDir())
            {
                //ProgramData目录缺少权限，已改用“我的文档”目录。服务维护流程已伴随关闭！
                Must_Use_PersonalDir = true;
                new Log($"      * 已采用备用权限策略！", ConsoleColor.DarkMagenta);
            }

            new Log($"     √ 已通过 {AppAttribute.AppName} 权限检查。", ConsoleColor.DarkGreen);
        }
    }
}
