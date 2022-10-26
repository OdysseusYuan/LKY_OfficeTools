/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeDownload.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.SDK.Aria2c;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// Office 下载类库
    /// </summary>
    internal class Lib_OfficeDownload
    {
        /// <summary>
        /// 下载文件列表
        /// </summary>
        static List<string> down_list = null;

        /// <summary>
        /// 重载实现下载
        /// </summary>
        internal Lib_OfficeDownload()
        {
            FilesDownload();
        }

        /// <summary>
        /// 下载所有文件（Aria2c）
        /// </summary>
        internal static void FilesDownload()
        {
            //获取下载列表
            down_list = Lib_OfficeInfo.Get_OfficeFileList();

            //定义下载目标地
            string save_to = Environment.CurrentDirectory + @"\Office\";

            //计划保存的地址
            List<string> save_files = new List<string>();

            //下载开始
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine($"\n------> 开始下载 Microsoft Office v{Lib_OfficeInfo.latest_version} 文件 ...");
            //延迟，让用户看到开始下载
            Thread.Sleep(1000);

            //轮询下载所有文件
            foreach (var a in down_list)
            {
                //根据官方目录，来调整下载保存位置
                string save_path = save_to + a.Substring(Lib_OfficeInfo.office_file_root_url.Length).Replace("/", "\\");

                //保存到List里面，用于后续检查
                save_files.Add(save_path);

                //遇到重复的文件可以断点续传
                Aria2c.DownFile(a, save_path);
            }

            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine($"\n------> 正在检查 Microsoft Office v{Lib_OfficeInfo.latest_version} 文件 ...");

            foreach (var b in save_files)
            {
                string aria_tmp_file = b + ".aria2c";

                //下载完成的标志：文件存在，且不存在临时文件
                if (File.Exists(b) && !File.Exists(aria_tmp_file))
                {
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine($"      > 正在检查 {new FileInfo(b).Name} ...");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"      > 文件 {new FileInfo(b).Name} 存在异常，重试中 ...");
                    FilesDownload();
                    return;
                }
            }

            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine($"-----√ Microsoft Office v{Lib_OfficeInfo.latest_version} 下载完成。\n");
        }

        ///迅雷下载方法存在BUG，暂停
        /*
        /// <summary>
        /// 下载单个文件（Thunder）
        /// </summary>
        /// <param name="url"></param>
        /// <param name="save_to"></param>
        internal static void DownFileByThunder(string url, string save_to)
        {
            if (down_list.Count > 0)
            {
                //Console.WriteLine(save_to);

                string file_path = new FileInfo(save_to).DirectoryName;     //保存的文件路径，不含文件名
                string filename = new FileInfo(save_to).Name;               //保存的文件名

                var manager = new DownloadManager(1, file_path);

                manager.CreateNewTask(url, new FileInfo(save_to).Name);

                Console.WriteLine($"下载至：{new FileInfo(save_to).DirectoryName}，文件名{filename}");

                manager.TaskDownload += (s, e) =>
                {
                    if (!(s is DownFileInfo info))
                    {
                        return;
                    }
                    Console.Write($"{filename} 已完成 {info.TaskInfo.Percent:0.0%}，速度：{info.TaskInfo.Speed / 1024 / 1024:0.0MB/S}\r");
                };

                manager.TaskCompleted += (s, e) =>
                {
                    if (!(s is DownFileInfo info))
                    {
                        return;
                    }
                    Console.WriteLine($"{info.FileName} 下载完成");
                };

                manager.StartAllTask();
            }
        }

        */
    }
}
