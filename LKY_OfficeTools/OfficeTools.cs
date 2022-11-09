/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : OfficeTools.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using LKY_OfficeTools.Lib;
using System;
using System.Reflection;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools
{
    internal class OfficeTools
    {
        static void Main(string[] args)
        {
            //命令行检测
            string arg = null;
            foreach (var now_arg in args)
            {
                arg += now_arg + ";";
            }

            Entry(arg);
        }

        /// <summary>
        /// 函数入口
        /// </summary>
        private static void Entry(string args = null)
        {
            //欢迎话术
            Version version = Assembly.GetExecutingAssembly().GetName().Version;

            //设置标题
            Console.Title = $"LKY Office Tools v{version}";

            //清理冗余信息
            Log.Clean();

            //数字签名证书检查
            new Lib_AppSignCert();

            //Header
            new Log($"LKY Office Tools [版本 {version}]\n" +
                $"版权所有（C）LiuKaiyuan (Odysseus.Yuan)。保留所有权利。\n\n" +
                $"探讨 {Console.Title} 相关内容，可发送邮件至：liukaiyuan@sjtu.edu.cn", ConsoleColor.Gray);

            //确认系统情况
            if (int.Parse(Com_SystemOS.OSVersion.GetBuildNumber()) < 15063)
            {
                //小于 Win10 1703 的操作系统，激活存在失败问题
                new Log($"\n     × 请将当前操作系统升级至 Windows 10 (1703) 或其以上版本，否则 Office 无法进行正版激活！", ConsoleColor.DarkRed);

                //退出机制
                QuitMsg();

                return;
            }

            //确认联网情况
            if (!Com_NetworkOS.Check.IsConnected)
            {
                new Log($"\n     × 请确保当前电脑可正常访问互联网！", ConsoleColor.DarkRed);

                //退出机制
                QuitMsg();

                return;
            }

            //根据命令行判断是否等待用户
            bool isContinue = false;
            if (!string.IsNullOrEmpty(args) && args.Contains("/none_welcome_confirm"))
            {
                isContinue = true;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write("\n请按 回车键 开始部署 ...");
                isContinue = (Console.ReadKey().Key == ConsoleKey.Enter);
            }
            
            if (isContinue)
            {
                //权限检查
                Com_PrivilegeOS.PrivilegeAttention();

                //更新检查
                Lib_AppUpdate.Check_Latest_Version();

                //继续
                new Lib_OfficeInstall();

                //日志回收
                Lib_AppCount.PostInfo.Finish();

                //退出机制
                QuitMsg();
            }
            else
            {
                //日志回收
                Lib_AppCount.PostInfo.Finish();

                return;
            }
        }

        private static void QuitMsg()
        {
            //退出机制
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("\n请按任意键退出 ...");
            Console.ReadKey();
        }
    }
}
