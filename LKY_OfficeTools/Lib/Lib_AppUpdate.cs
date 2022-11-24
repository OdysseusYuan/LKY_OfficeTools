/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_SelfUpdate.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading;
using static LKY_OfficeTools.Common.Com_FileOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppReport;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 用于检查更新的类库
    /// </summary>
    internal class Lib_AppUpdate
    {
#if (!DEBUG)
        //release模式地址
        /// <summary>
        /// 获取更新info时的访问地址
        /// </summary>
        internal const string update_json_url = "https://gitee.com/OdysseusYuan/LKY_OfficeTools/releases/download/AppInfo/LKY_OfficeTools_AppInfo.json";
#else
        //debug模式地址，test.json
        /// <summary>
        /// 获取更新info时的访问地址
        /// </summary>
        internal const string update_json_url = "https://gitee.com/OdysseusYuan/LKY_OfficeTools/releases/download/AppInfo/test.json";
#endif

        /// <summary>
        /// 获取到的最新信息
        /// </summary>
        internal static string latest_info { get; set; }

        /// <summary>
        /// 检查自身最新版本
        /// </summary>
        /// <returns></returns>
        internal static bool Check()
        {
            try
            {
                new Log($"\n------> 正在进行 {Console.Title} 初始化检查 ...", ConsoleColor.DarkCyan);

                //当更新完成自重启时，自动删除 .old 的文件
                ScanFiles oldFiles = new ScanFiles();
                oldFiles.GetFilesByExtension(AppPath.ExecuteDir, ".old");
                ///无 old 文件时，自动跳过
                if (oldFiles.FilesList != null && oldFiles.FilesList.Count > 0)
                {
                    foreach (var now_file in oldFiles.FilesList)
                    {
                        File.Delete(now_file);
                    }
                }

                new Log($"     >> 初始化完成 {new Random().Next(1, 10)}% ...", ConsoleColor.DarkYellow);

                //获取版本信息
                latest_info = Com_WebOS.Visit_WebClient(update_json_url);

                //截取获得最新版本和下载地址
                string latest_ver = Com_TextOS.GetCenterText(latest_info, "\"Latest_Version\": \"", "\"");
                string latest_down_url = Com_TextOS.GetCenterText(latest_info, "\"Latest_Version_Update_Url\": \"", "\"");

                new Log($"     >> 初始化完成 {new Random().Next(11, 30)}% ...", ConsoleColor.DarkYellow);

                //启动打点
                Pointing(ProcessStage.Starting, true);

                new Log($"     >> 初始化完成 {new Random().Next(91, 100)}% ...", ConsoleColor.DarkYellow);

                new Log($"     √ 已完成 {Console.Title} v{latest_ver} 初始化检查。", ConsoleColor.DarkGreen);

                string now_ver = AppAttribute.AppVersion;
                if (new Version(latest_ver) > new Version(now_ver))
                {
                    //发现新版本
                    new Log($"\n------> 正在更新 {Console.Title} ...", ConsoleColor.DarkCyan);

                    //下载文件
                    string save_to = AppPath.Documents.Documents_Root + @"\Update\" + $"{Console.Title}_updateto_{latest_ver}.zip";
                    int down_result = Lib_Aria2c.DownFile(latest_down_url, save_to);

                    //下载不成功时，抛出
                    if (down_result != 1)
                    {
                        throw new Exception();
                    }

                    //解压文件
                    string extra_to = Path.GetDirectoryName(save_to) + "\\" + $"{Console.Title}_update_{latest_ver}";
                    ///如果目录存在，先清空下目标文件夹，删除子目录、子文件等
                    if (Directory.Exists(extra_to))
                    {
                        Directory.Delete(extra_to, true);
                    }
                    ZipFile.ExtractToDirectory(save_to, extra_to);
                    ///删除下载的zip
                    File.Delete(save_to);

                    //扫描文件
                    ScanFiles new_files = new ScanFiles();
                    new_files.GetFilesByExtension(extra_to);
                    ///可更新文件为空，跳过更新
                    if (new_files.FilesList == null)
                    {
                        throw new Exception();
                    }

                    //获得自身主程序路径
                    string self_RunPath = AppPath.Executer;

                    //删除旧的文件
                    ScanFiles old_files = new ScanFiles();
                    old_files.GetFilesByExtension(extra_to);
                    if (old_files.FilesList != null && old_files.FilesList.Count > 0)
                    {
                        foreach (var now_file in old_files.FilesList)
                        {
                            //除了自身exe外，删除全部旧的文件
                            if (now_file != self_RunPath)
                            {
                                try
                                {
                                    File.Delete(now_file);
                                }
                                catch (Exception Ex)
                                {
                                    new Log(Ex.ToString());
                                }
                            }
                        }
                    }

                    //复制新文件
                    foreach (var now_file in new_files.FilesList)
                    {
                        //获得文件相对路径
                        string file_relative_path = now_file.Replace(extra_to, "\\");
                        //合成移动路径
                        string move_to = AppPath.ExecuteDir + file_relative_path;
                        //替换自身时，先move为别的文件，然后再替换
                        if (new FileInfo(move_to).FullName == self_RunPath)
                        {
                            File.Move(self_RunPath, self_RunPath + ".old");
                        }

                        //增加目录创建，否则目标文件拷贝将会失败
                        string move_dir = new FileInfo(move_to).DirectoryName;
                        Directory.CreateDirectory(move_dir);                    //目录已经存在时，重复创建，不会引发异常

                        File.Copy(now_file, move_to, true);
                    }

                    //更新后，删除更新目录
                    if (Directory.Exists(extra_to))
                    {
                        Directory.Delete(extra_to, true);
                    }

                    //重启自身完成更新
                    new Log($"\n     √ 已更新至 {Console.Title} v{latest_ver} 版本，程序即将自动重启，请稍候。", ConsoleColor.DarkGreen);

                    Thread.Sleep(5000);

                    //启动实例
                    Process p = new Process();
                    p.StartInfo.FileName = AppPath.Executer;             //需要启动的程序名       
                    p.StartInfo.Arguments = "/none_welcome_confirm";         //启动参数
                    p.Start();
                    //关闭当前实例
                    Process.GetCurrentProcess().Kill();
                }

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                new Log($"      * 暂时跳过更新检查！", ConsoleColor.DarkMagenta);
                return false;
            }

        }
    }
}
