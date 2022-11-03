/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Threading;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_OfficeInfo
    {
        /// <summary>
        /// 查看 Office 网络的版本类库
        /// </summary>
        internal class OfficeNetVersion
        {
            /// <summary>
            /// 最新版 Office 版本信息
            /// </summary>
            internal static Version latest_version = null;

            private const string office_info_url = "https://config.office.com/api/filelist/channels";

            public static string office_file_root_url;

            /// <summary>
            /// 获取最新版本的 Office 函数
            /// </summary>
            /// <returns></returns>
            internal static Version Get_OfficeVersion()
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.DarkCyan;
                    Console.WriteLine("\n------> 正在获取最新 Microsoft Office 版本 ...");

                    //获取频道信息       
                    string office_info = Com_WebOS.Visit_WebRequest(office_info_url);

                    if (!string.IsNullOrEmpty(office_info))
                    {
                        //获取版本信息
                        string latest_info = Com_TextOS.GetCenterText(office_info, "\"PerpetualVL2021\",", "name");     //获取 2021 LTSC
                        latest_version = new Version(Com_TextOS.GetCenterText(latest_info, "latestUpdateVersion\":\"", "\"},"));              //获取版本号

                        //赋值对应的下载地址
                        office_file_root_url = Com_TextOS.GetCenterText(latest_info, "baseUrl\":\"", "\"");              //获取url

                        //Console.WriteLine(office_file_root_url);

                        return latest_version;
                    }
                    else
                    { return null; }
                }

                catch (WebException webEx)
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine(webEx.Message.ToString());
                    return null;
                }

            }

            /// <summary>
            /// 获取文件下载列表
            /// </summary>
            /// <returns></returns>
            internal static List<string> Get_OfficeFileList()
            {
                try
                {
                    Version version_info = Get_OfficeVersion();      //获取版本

                    if (version_info == null || string.IsNullOrEmpty(office_file_root_url))     //下载根地址为空时，视为失败
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine("     × 最新版本获取失败，请稍后重试！");
                        return null;
                    }

                    string ver = version_info.ToString();

                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine($"     √ 已完成，最新版：{ver}。");

                    //延迟，让用户看清版本号
                    Thread.Sleep(500);


                    //获取文件列表
                    List<string> file_list = new List<string>();
                    office_file_root_url += "/office/data";
                    ///获取当前系统位数
                    int sys_bit;
                    if (Environment.Is64BitOperatingSystem)
                    {
                        sys_bit = 64;
                    }
                    else
                    {
                        sys_bit = 32;
                    }

                    //x32系统也需要下载 64 的 i64****.cab文件，但必须放在 {ver} 目录下。
                    file_list.Add($"{office_file_root_url}/{ver}/i640.cab");
                    file_list.Add($"{office_file_root_url}/{ver}/i642052.cab");

                    switch (sys_bit)
                    {
                        case 64:
                            file_list.Add($"{office_file_root_url}/{ver}/stream.x{sys_bit}.x-none.dat");
                            file_list.Add($"{office_file_root_url}/{ver}/stream.x{sys_bit}.zh-cn.dat");
                            break;
                        case 32:
                            file_list.Add($"{office_file_root_url}/{ver}/stream.x86.x-none.dat");
                            file_list.Add($"{office_file_root_url}/{ver}/stream.x86.zh-cn.dat");
                            file_list.Add($"{office_file_root_url}/{ver}/i{sys_bit}0.cab");
                            file_list.Add($"{office_file_root_url}/{ver}/i{sys_bit}2052.cab");
                            break;
                    }

                    //按版本下载其余文件
                    file_list.Add($"{office_file_root_url}/v{sys_bit}.cab");
                    file_list.Add($"{office_file_root_url}/v{sys_bit}_{ver}.cab");
                    file_list.Add($"{office_file_root_url}/{ver}/s{sys_bit}0.cab");
                    file_list.Add($"{office_file_root_url}/{ver}/s{sys_bit}2052.cab");

                    /*
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine($"\n------> 预计下载文件：");
                    foreach (var a in file_list)
                    {
                        Console.WriteLine($"      > {a}");
                    }
                    Console.Read();
                    */

                    return file_list;
                }

                catch (WebException webEx)
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine(webEx.Message.ToString());
                    return null;
                }

            }
        }

        /// <summary>
        /// 查看本地 Office 安装情况 类库
        /// </summary>
        internal class OfficeLocalInstall
        {
            /// <summary>
            /// Office 本地安装情况
            /// </summary>
            internal enum State
            {
                /// <summary>
                /// 本机未安装任何 Office 或 没有使用本软件安装 Office ，通过是否存在 ClickToRun 项判断
                /// </summary>
                Nothing,

                /// <summary>
                /// 安装了Office，但不是最新版
                /// </summary>
                VersionDiff,

                /// <summary>
                /// 已安装了最新版 Office
                /// </summary>
                Installed
            }

            internal static State GetState()
            {
                //检查注册表，判断安装是否成功
                RegistryKey HKLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                        Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);      //判断操作系统版本（64位\32位）打开注册表项，不然 x86编译的本程序 读取 x64的程序会出现无法读取 已经存在于注册表 中的数据

                RegistryKey office_reg = HKLM.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");

                if (office_reg == null)
                {
                    //找不到 ClickToRun 注册表
                    return State.Nothing;
                }
                else
                {
                    object office_InstallVer = office_reg.GetValue("VersionToReport");
                    if (office_InstallVer != null && office_InstallVer.ToString() == OfficeNetVersion.latest_version.ToString())      //必须先判断不为null，否则会抛出异常
                    {
                        //一切正常
                        return State.Installed;
                    }
                    else
                    {
                        //版本号和一开始下载的版本号不一致
                        return State.VersionDiff;
                    }
                }
            }
        }
    }
}
