/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_SelfUpdate.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using LKY_OfficeTools.SDK.Aria2c;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using static LKY_OfficeTools.Common.Com_FileOS;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 用于检查更新的类库
    /// </summary>
    internal class Lib_SelfUpdate
    {
        /// <summary>
        /// 检查自身最新版本
        /// </summary>
        /// <returns></returns>
        internal static bool Check_Latest_Version()
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine($"\n------> 正在进行 {Console.Title} 初始化检查 ...");

                //当更新完成自重启时，自动删除 .old 的文件
                ScanFiles oldFiles = new ScanFiles();
                oldFiles.GetFilesByExtension(Environment.CurrentDirectory, ".old");
                ///无 old 文件时，自动跳过
                if (oldFiles.FilesList != null)
                {
                    foreach (var now_file in oldFiles.FilesList)
                    {
                        File.Delete(now_file);
                    }
                }

                string check_url = "https://gitee.com/OdysseusYuan/LKY_OfficeTools/releases/download/AppInfo/LKY_OfficeTools_AppInfo.json";

                //获取 github 版本信息
                WebClient MyWebClient = new WebClient();
                MyWebClient.Credentials = CredentialCache.DefaultCredentials;       //获取或设置用于向Internet资源的请求进行身份验证的网络凭据
                Byte[] pageData = MyWebClient.DownloadData(check_url);        //从指定网站下载数据
                                                                              //string pageHtml = Encoding.Default.GetString(pageData);             //如果获取网站页面采用的是GB2312，则使用这句            
                string latest_info = Encoding.UTF8.GetString(pageData);             //如果获取网站页面采用的是UTF-8，则使用这句

                //截取获得最新版本和下载地址
                string latest_ver = Com_TextOS.GetCenterText(latest_info, "\"Latest_Version\": \"", "\"");
                string latest_down_url = Com_TextOS.GetCenterText(latest_info, "\"Latest_Version_Update_Url\": \"", "\"");

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ 已完成 {Console.Title} v{latest_ver} 初始化检查。");

                string now_ver = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                if (new Version(latest_ver) > new Version(now_ver))
                {
                    //发现新版本
                    Console.ForegroundColor = ConsoleColor.DarkCyan;
                    Console.WriteLine($"\n------> 正在更新 {Console.Title} ...");

                    //下载文件
                    string save_to = Environment.CurrentDirectory + @"\Update\" + $"{Console.Title}_updateto_{latest_ver}.zip";
                    Aria2c.DownFile(latest_down_url, save_to);

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
                    ScanFiles scanFiles = new ScanFiles();
                    scanFiles.GetFilesByExtension(extra_to);
                    ///可更新文件为空，跳过更新
                    if (scanFiles.FilesList == null)
                    {
                        throw new Exception();
                    }

                    //复制文件
                    ///获得自身主程序路径
                    string self_RunPath = new FileInfo(Process.GetCurrentProcess().MainModule.FileName).FullName;
                    foreach (var now_file in scanFiles.FilesList)
                    {
                        //获得文件相对路径
                        string file_relative_path = now_file.Replace(extra_to, "\\");
                        //合成移动路径
                        string move_to = Environment.CurrentDirectory + file_relative_path;
                        //替换自身时，先move为别的文件，然后再替换
                        if (new FileInfo(move_to).FullName == self_RunPath)
                        {
                            File.Move(self_RunPath, self_RunPath + ".old");
                        }
                        File.Copy(now_file, move_to, true);
                    }

                    //更新后，删除更新目录
                    if (Directory.Exists(extra_to))
                    {
                        Directory.Delete(extra_to, true);
                    }

                    //重启自身完成更新
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine($"\n     √ 已更新至 {Console.Title} v{latest_ver} 版本，程序即将自动重启，请稍候。");

                    Thread.Sleep(3000);

                    //启动实例
                    Process.Start(Process.GetCurrentProcess().MainModule.FileName);
                    //关闭当前实例
                    Process.GetCurrentProcess().Kill();
                }

                return true;
            }
            catch /*(Exception Ex)*/
            {
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"      * 暂时跳过更新检查！");
                //Console.WriteLine(Ex.Message.ToString());
                return false;
            }

        }
    }
}
