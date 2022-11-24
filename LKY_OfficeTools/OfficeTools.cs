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
using System.Text;
using static LKY_OfficeTools.Lib.Lib_AppCommand;
using static LKY_OfficeTools.Lib.Lib_AppInfo.App.State;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppReport;

namespace LKY_OfficeTools
{
    internal class OfficeTools
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.GetEncoding("gbk");       //设定编码，解决英文系统乱码问题

            //中断检测
            Close.SetConsoleCtrlHandler(Close.newDelegate, true);

            //命令行检测
            new Lib_AppCommand(args);

            //启动
            Entry();
        }

        /// <summary>
        /// 函数入口
        /// </summary>
        private static void Entry()
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

                //退出提示
                KeyMsg.Quit();

                return;
            }

            //确认联网情况
            if (!Com_NetworkOS.Check.IsConnected)
            {
                new Log($"\n     × 请确保当前电脑可正常访问互联网！", ConsoleColor.DarkRed);

                //退出提示
                KeyMsg.Quit();

                return;
            }

            //根据命令行判断是否等待用户，没有标记时，需要人工来决定
            bool isContinue = true;
            if (!AppCommandFlag.HasFlag(ArgsFlag.None_Welcome_Confirm))
            {
                isContinue = (KeyMsg.Confirm());
            }

            if (isContinue)
            {
                //权限检查
                Com_PrivilegeOS.PrivilegeAttention();

                //SDK初始化
                Lib_AppSdk.initial();

                //更新检查
                Lib_AppUpdate.Check();

                //继续
                new Lib_OfficeInstall();

                Pointing(Current_StageType, true);    //回收

                //展现一条龙服务的结论
                if (Current_StageType == ProcessStage.Finish_Fail)
                {
                    new Log($"\n     × 当前部署存在失败环节，您可在稍后重试运行！", ConsoleColor.DarkRed);
                }

                //退出提示
                KeyMsg.Quit();
            }
            else
            {
                new Log($"\n     × 您未按 回车键，软件已停止部署！", ConsoleColor.DarkRed);
                Pointing(ProcessStage.Finish_Fail);  //回收
                return;
            }
        }


    }
}
