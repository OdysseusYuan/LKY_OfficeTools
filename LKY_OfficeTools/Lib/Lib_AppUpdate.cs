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
        /// <summary>
        /// 检查自身最新版本
        /// </summary>
        /// <returns></returns>
        internal static bool Check()
        {
            try
            {
                new Log($"\n------> 正在进行 {AppAttribute.AppName} 初始化检查 ...", ConsoleColor.DarkCyan);

                //当更新完成自重启时，自动删除 Update Trash 目录中的 .old 文件
                ScanFiles oldFiles = new ScanFiles();
                oldFiles.GetFilesByExtension(AppPath.Documents.Update.UpdateTrash, ".old");
                ///无 old 文件时，自动跳过
                if (oldFiles.FilesList != null && oldFiles.FilesList.Count > 0)
                {
                    foreach (var now_file in oldFiles.FilesList)
                    {
                        //使用 try catch 模式。在服务模式下，更新完成后，old 文件可能还会被服务占用，此时跳过该文件，避免占用报错，添加 try 模式。
                        try
                        {
                            File.Delete(now_file);
                        }
                        catch { }
                    }
                }

                new Log($"     >> 初始化完成 {new Random().Next(1, 10)}% ...", ConsoleColor.DarkYellow);

                //截取获得最新版本和下载地址
                string latest_ver = Com_TextOS.GetCenterText(AppJson.Info, "\"Latest_Version\": \"", "\"");
                string latest_down_url = Com_TextOS.GetCenterText(AppJson.Info, "\"Latest_Version_Update_Url\": \"", "\"");

                new Log($"     >> 初始化完成 {new Random().Next(11, 30)}% ...", ConsoleColor.DarkYellow);

                //启动打点
                Pointing(ProcessStage.Starting, true);

                new Log($"     >> 初始化完成 {new Random().Next(91, 100)}% ...", ConsoleColor.DarkYellow);

                new Log($"     √ 已完成 {AppAttribute.AppName} 初始化检查。", ConsoleColor.DarkGreen);

                string now_ver = AppAttribute.AppVersion;
                if (new Version(latest_ver) > new Version(now_ver))
                {
                    //发现新版本
                    new Log($"\n------> 正在更新 {AppAttribute.AppName} 至 v{latest_ver} 版本 ...", ConsoleColor.DarkCyan);
                                       
                    new Log($"\n     >> 下载 v{latest_ver} 更新包中 ...", ConsoleColor.DarkYellow);

                    //下载文件
                    string save_to = AppPath.Documents.Update.Update_Root + $"\\v{latest_ver}.zip";

                    //下载前先删除旧的文件（禁止续传），否则一旦意外中断，再次启动下载将出现异常
                    try
                    {
                        //删除主体文件
                        if (File.Exists(save_to))
                        { 
                            File.Delete(save_to);
                        }

                        //删除主体对应的描述文件
                        string des_file = save_to + ".aria2";
                        if (File.Exists(des_file))
                        {
                            File.Delete(des_file);
                        }
                    }
                    catch 
                    { 
                        //仅用于日志记录
                        new Log($"清理冗余更新包文件失败，后续下载可能会出现异常！"); 
                    }

                    //开始下载
                    int down_result = Lib_Aria2c.DownFile(latest_down_url, save_to);

                    //下载不成功时，抛出
                    if (down_result != 1)
                    {
                        throw new Exception();
                    }

                    new Log($"\n     >> 更新 v{latest_ver} 文件中 ...", ConsoleColor.DarkYellow);

                    //解压文件
                    string extra_to = Path.GetDirectoryName(save_to) + "\\" + $"v{latest_ver}";
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

                    /* ---------> 暂停删除旧文件。当用户把本文件放在 桌面 或有其他文件的 文件夹中时，会导致资料连带删除！
                    //删除旧的文件
                    ScanFiles old_files = new ScanFiles();
                    old_files.GetFilesByExtension(AppPath.ExecuteDir);
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
                    */

                    //复制新文件
                    foreach (var now_file in new_files.FilesList)
                    {
                        //获得文件相对路径
                        string file_relative_path = now_file.Replace(extra_to, "\\");
                        //合成移动路径
                        string move_to = AppPath.ExecuteDir + file_relative_path;

                        //如果新文件在现有路径中存在，则将旧的文件先 move 到 Trash 目录，解决文件被占用的情况
                        if (File.Exists(move_to))
                        {
                            string dest_filepath = AppPath.Documents.Update.UpdateTrash + $"\\{DateTime.Now.ToFileTime()}.old";
                            Directory.CreateDirectory(Path.GetDirectoryName(dest_filepath));        //创建目录
                            File.Move(move_to, dest_filepath);
                        }

                        //增加目录创建，否则目标文件拷贝将会失败
                        Directory.CreateDirectory(Path.GetDirectoryName(move_to));                    //目录已经存在时，重复创建，不会引发异常

                        //拷贝、覆盖新文件
                        File.Copy(now_file, move_to, true);
                    }

                    //若旧的 exe 文件名和默认 exe 文件名不一致时，将自身 exe 文件 move 到 Trash 目录。
                    //解决用户修改旧 exe 文件名，复制新文件后，会出现两个 exe 的情况。
                    if (AppPath.Executer != (AppPath.ExecuteDir + $"\\{AppAttribute.AppFilename}"))
                    {
                        string exe_moveto = AppPath.Documents.Update.UpdateTrash + $"\\{DateTime.Now.ToFileTime()}.old";
                        Directory.CreateDirectory(Path.GetDirectoryName(exe_moveto));        //创建目录
                        File.Move(AppPath.Executer, exe_moveto);
                    }

                    //更新后，删除更新目录
                    if (Directory.Exists(extra_to))
                    {
                        Directory.Delete(extra_to, true);
                    }

                    //重启自身完成更新
                    new Log($"\n     √ 已更新至 {AppAttribute.AppName} v{latest_ver} 版本，程序即将自动重启，请稍候。", ConsoleColor.DarkGreen);
                    Thread.Sleep(5000);

                    //重启进程
                    RestartProcess();
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

        /// <summary>
        /// 更新完成，重启进程或服务，完成新的部署
        /// </summary>
        /// <returns></returns>
        static bool RestartProcess()
        {
            try
            {
                /* 
                 * 【设计背景】
                 *  1、自动更新在服务模式下的 confirm 卡进程问题。
                 *  2、服务自行升级时，用户手动停止服务，无法关闭更新后文件再运行的进程ID。
                 * 
                 * 【情况判断】
                 *  一、已安装服务
                 *  （1）服务文件存在：
                 *      A. 自定义位置：服务文件没有被更新，即使执行 passive 指令，也是显示的运行，故，进程重启
                 *      B. 服务位置：文件已被更新，且通常为 passive 模式，需要继承 passive 指令，否则会卡在 confirm 环节，故，重启服务
                 *  （2）服务文件不存在：服务无法执行 passvive，因此不会卡在 confirm 环节，故，进程重启
                 *  二、没有安装服务：只能手动模式，即使执行 passive 指令，也是显示的运行，不会卡进程。故，进程重启
                 * 
                 * 【结论】
                 *  综上，当且仅当 服务已安装 & 在服务目录运行（服务文件存在） 时，更新后重启服务，除此之外，全部 只重启进程
                */

                if (//重启服务条件
                    Com_ServiceOS.Query.IsCreated(AppAttribute.ServiceName) &&              //服务已被创建
                    AppPath.Executer == AppPath.Documents.Services.ServiceAutorun_Exe       //当前运行位置为服务文件位置（满足该条件，服务文件天然存在）
                    )
                {
                    //满足服务重启的条件，重启服务，运行新的exe
                    Lib_AppServiceConfig.RestartSelf();                                     //重启服务后，软件已经是最新版exe，并带着 passive 指令运行
                }
                else
                {
                    //未安装服务 OR 在非服务文件路径运行，运行进程
                    /*
                     * 当用户手动修改了升级旧exe的名字，这时 Executer 的名字 和 更新后文件exe的名字不一致。使用 Executer 重启，会导致重启进程失败。
                     * 因此，应按照默认文件名，重启进程。这要求后续每个升级版本的默认主执行文件 exe 文件名，不得发生改变。
                     */

                    //定义路径
                    string run_path = AppPath.ExecuteDir + $"\\{AppAttribute.AppFilename}";

                    //启动实例
                    Process p = new Process();
                    p.StartInfo.FileName = run_path;                         //需要启动的程序名      
                    p.StartInfo.Arguments = "/none_welcome_confirm";         //启动参数
                    p.Start();

                    /*
                     * 暂时不使用 Process.StartInfo.UseShellExecute = false 模式，否则在执行时，会偶发性出现：应用程序无法正常启动(0xc0000142) 的错误
                    Com_ExeOS.Run.Exe(run_path, "/none_welcome_confirm", false);    //二者 一般为 手动模式运行，无需 passive，默认使用 跳过欢迎确认 指令
                    */
                }

                //无论何种模式，均要关闭当前旧的实例
                Process.GetCurrentProcess().Kill();

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                new Log($"      * 暂时跳过更新步骤！", ConsoleColor.DarkMagenta);
                return false;
            }
        }
    }
}
