/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppService.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
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

        /// <summary>
        /// 启动服务后事件
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            try
            {
                //每次启动服务时，自动删除 Services Trash 目录中的 .old 文件
                ScanFiles oldFiles = new ScanFiles();
                oldFiles.GetFilesByExtension(AppPath.Documents.Services.ServicesTrash, ".old");
                ///无 old 文件时，自动跳过
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
                ///创建服务目录
                Directory.CreateDirectory(AppPath.Documents.Services.Services_Root);
                ///保存进程ID到文件（覆盖模式）
                string info_file = AppPath.Documents.Services.PassiveProcessInfo;
                File.WriteAllText(info_file, process_info.Id.ToString());
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }

        /// <summary>
        /// 停止服务后事件。
        /// 此处不能调用任何 本服务内部变量 的行为。否则会引发 1601 错误。
        /// </summary>
        protected override void OnStop()
        {
            try
            {
                //服务停止时，如果有正在运行的 passive 进程，立即结束。
                ///判断进程信息文件是否存在
                string info_path = AppPath.Documents.Services.PassiveProcessInfo;
                if (File.Exists(info_path))
                {
                    ///读取进程ID
                    string info = File.ReadAllText(info_path);
                    ///判断是否为空
                    if (!string.IsNullOrWhiteSpace(info))
                    {
                        ///不为空时，尝试结束进程
                        Com_ExeOS.KillExe.ByProcessID(int.Parse(info), Com_ExeOS.KillExe.KillMode.Try_Friendly, true);       //尝试友好的结束进程
                    }

                    ///完成后，删除文件
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
