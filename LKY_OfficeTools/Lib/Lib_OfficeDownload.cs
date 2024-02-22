/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 LiuKaiyuan. All rights reserved.
 *      
 *      FileName : Lib_OfficeDownload.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_OfficeDownload
    {
        internal static int StartDownload()
        {
            try
            {
                //定义下载目标地。文件必须位于 \Office\Data\下。ODT安装必须在 Office 上一级目录上执行。
                string save_to = AppPath.ExecuteDir + @"\Office\Data\";

                //计划保存的地址
                List<string> save_files = new List<string>();

                //下载开始
                new Log($"\n------> 开始下载 Office v{OfficeNetInfo.OfficeLatestVersion} 文件 ...", ConsoleColor.DarkCyan);
                //延迟，让用户看到开始下载
                Thread.Sleep(1000);

                //轮询下载所有文件
                foreach (var a in OfficeNetInfo.OfficeFileList)
                {
                    //根据官方目录，来调整下载保存位置
                    string save_path = save_to + a.Substring(OfficeNetInfo.OfficeUrlRoot.Length).Replace("/", "\\");

                    //保存到List里面，用于后续检查
                    save_files.Add(save_path);

                    new Log($"\n     >> 下载 {new FileInfo(save_path).Name} 文件中，请稍候 ...", ConsoleColor.DarkYellow);

                    //遇到重复的文件可以断点续传
                    int down_result = Lib_Aria2c.DownFile(a, save_path);
                    if (down_result != 1)
                    {
                        //如果因为核心下载exe丢失，导致下载失败，直接中止
                        throw new Exception($"下载 {a} 异常！");
                    }

                    //如果用户中断了下载，则直接跳出
                    if (Lib_AppState.Current_StageType == Lib_AppState.ProcessStage.Interrupt)
                    {
                        return -1;
                    }

                    new Log($"     √ 已下载 {new FileInfo(save_path).Name} 文件。", ConsoleColor.DarkGreen);
                }

                new Log($"\n------> 正在检查 Office v{OfficeNetInfo.OfficeLatestVersion} 文件 ...", ConsoleColor.DarkCyan);

                foreach (var b in save_files)
                {
                    string aria_tmp_file = b + ".aria2c";

                    //下载完成的标志：文件存在，且不存在临时文件
                    if (File.Exists(b) && !File.Exists(aria_tmp_file))
                    {
                        new Log($"     >> 检查 {new FileInfo(b).Name} 文件中 ...", ConsoleColor.DarkYellow);
                    }
                    else
                    {
                        new Log($"     >> 文件 {new FileInfo(b).Name} 存在异常，重试中 ...", ConsoleColor.DarkRed);
                        return StartDownload();
                    }
                }

                new Log($"     √ 已完成 Office v{OfficeNetInfo.OfficeLatestVersion} 下载。", ConsoleColor.DarkGreen);

                return 1;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return 0;
            }
        }
    }
}
