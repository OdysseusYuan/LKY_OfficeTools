/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2023 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : OfficeTools.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using LKY_OfficeTools.Lib;
using System;
using System.Text;
using static LKY_OfficeTools.Lib.Lib_AppClosing;
using static LKY_OfficeTools.Lib.Lib_AppCommand;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppReport;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools
{
    internal class OfficeTools
    {
        static void Main(string[] args)
        {
            //命令行检测
            new Lib_AppCommand(args);

            //设定编码，解决英文系统乱码问题
            Console.OutputEncoding = Encoding.GetEncoding("gbk");       //必须在启动模式判断之后再进行本函数执行。否则服务模式将引发 1053 错误！

            //中断检测
            CloseWindow.SetConsoleCtrlHandler(CloseWindow.newDelegate, true);

            //启动
            Entry();
        }

        /// <summary>
        /// 函数入口
        /// </summary>
        private static void Entry()
        {
            //欢迎话术
            ///设置标题
            Console.Title = $"{AppAttribute.AppName} v{AppAttribute.AppVersion}";
            ///Header
            new Log($"{AppAttribute.AppName} [版本 {AppAttribute.AppVersion}]\n" +
                $"版权所有（C）LiuKaiyuan (Odysseus.Yuan)。保留所有权利。\n\n" +
                $"探讨 {Console.Title} 相关内容，可发送邮件至：liukaiyuan@sjtu.edu.cn", ConsoleColor.Gray);

            //清理冗余信息
            Log.Clean();

            //数字签名证书检查
            new Lib_AppSignCert();

            //确认系统情况
            if (int.Parse(Com_SystemOS.OSVersion.GetBuildNumber()) < 15063)
            {
                //小于 Win10 1703 的操作系统，激活存在失败问题
                new Log($"\n     × 请将当前操作系统升级至 Windows 10 (1703) 或其以上版本，否则 Office 无法进行正版激活！", ConsoleColor.DarkRed);

                //退出提示
                KeyMsg.Quit(-4);

                return;
            }

            //确认联网情况
            if (!Com_NetworkOS.Check.IsConnected)
            {
                new Log($"\n     × 请确保当前电脑可正常访问互联网！", ConsoleColor.DarkRed);

                //退出提示
                KeyMsg.Quit(-5);

                return;
            }

            //根据命令行判断是否等待用户，没有标记时，执行倒计时等待
            if (!AppCommandFlag.HasFlag(ArgsFlag.None_Welcome_Confirm))
            {
                KeyMsg.DoByTime($"部署", 5);
            }

            //权限检查
            Com_PrivilegeOS.PrivilegeAttention();

            //SDK初始化
            Lib_AppSdk.Initial();

            //更新检查
            Lib_AppUpdate.Check();

            //继续
            new Lib_OfficeInstall();

            //部署成功时，提示是否配置为服务
            if (Current_StageType == ProcessStage.Finish_Success)
            {
                //被强制使用“我的文档”目录时，往往因为权限问题才使用，此时禁用服务配置
                if (!Must_Use_PersonalDir)
                {
                    //配置服务
                    Lib_AppServiceConfig.Setup();
                }

                //成功，配置后回收
                Pointing(Current_StageType, true);

                //结论
                new Log($"\n     √ 您已成功完成 {AppAttribute.AppName} 所有流程，感谢您的使用。", ConsoleColor.DarkGreen);
            }
            //部署失败，提示错误信息
            else if (Current_StageType == ProcessStage.Finish_Fail)
            {
                //失败，先回收，再显示结论
                Pointing(Current_StageType, true);

                //结论
                new Log($"\n     × 当前部署存在失败环节，您可在稍后重试运行！", ConsoleColor.DarkRed);
            }

            //退出提示
            KeyMsg.Quit(0);

        }


    }
}
