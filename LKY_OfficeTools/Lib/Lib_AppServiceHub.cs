/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 LiuKaiyuan. All rights reserved.
 *      
 *      FileName : Lib_AppService.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using static LKY_OfficeTools.Common.Com_FileOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    partial class Lib_AppServiceHub : ServiceBase
    {
        public Lib_AppServiceHub()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                //每次启动服务时，自动删除 Services Trash 目录中的 .old 文件
                ScanFiles oldFiles = new ScanFiles();
                oldFiles.GetFilesByExtension(AppPath.Documents.Services.ServicesTrash, ".old");
                if (oldFiles.FilesList != null && oldFiles.FilesList.Count > 0)
                {
                    foreach (var now_file in oldFiles.FilesList)
                    {
                        //使用 try catch 模式。以防异常。
                        try
                        {
                            File.Delete(now_file);
                        }
                        catch { }
                    }
                }

                //以无人值守的模式，隐式的运行本程序。
                Process process_info = new Process();
                Com_ExeOS.Run.Process(AppPath.Executer, "/passive", out process_info, false);      //异步运行

                //保存进程信息
                Directory.CreateDirectory(AppPath.Documents.Services.Services_Root);
                string info_file = AppPath.Documents.Services.PassiveProcessInfo;
                File.WriteAllText(info_file, process_info.Id.ToString());
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }

        protected override void OnStop()
        {
            try
            {
                //服务停止时，如果有正在运行的 passive 进程，立即结束。
                string info_path = AppPath.Documents.Services.PassiveProcessInfo;
                if (File.Exists(info_path))
                {
                    string info = File.ReadAllText(info_path);
                    if (!string.IsNullOrWhiteSpace(info))
                    {
                        Com_ExeOS.KillExe.ByProcessID(int.Parse(info), Com_ExeOS.KillExe.KillMode.Try_Friendly, true);       //尝试友好的结束进程
                    }

                    File.Delete(info_path);
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }
    }
}
