/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_OfficeInfo
    {
        /// <summary>
        /// Office 每一代发行的架构
        /// </summary>
        internal enum OfficeArchi
        {
            /// <summary>
            /// Office 2003及早期版本
            /// </summary>
            Office_Lower = 2003,

            /// <summary>
            /// Office 12.0
            /// </summary>
            Office_2007 = 2007,

            /// <summary>
            /// Office 14.0
            /// </summary>
            Office_2010 = 2010,

            /// <summary>
            /// Office 15.0
            /// </summary>
            Office_2013 = 2013,

            /// <summary>
            /// Office 16.0及其以上版本
            /// </summary>
            Office_ODT = 2016,
        }

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
            internal static Version GetOfficeVersion()
            {
                try
                {
                    new Log("\n------> 正在获取 最新 Office 版本 ...", ConsoleColor.DarkCyan);

                    //获取频道信息       
                    string office_info = Com_WebOS.Visit_WebClient(office_info_url);

                    if (!string.IsNullOrEmpty(office_info))
                    {
                        //获取版本信息
                        string latest_info = Com_TextOS.GetCenterText(office_info, "\"PerpetualVL2021\",", "name");     //获取 2021 LTSC
                        latest_version = new Version(Com_TextOS.GetCenterText(latest_info, "latestUpdateVersion\":\"", "\"},"));              //获取版本号

                        //赋值对应的下载地址
                        office_file_root_url = Com_TextOS.GetCenterText(latest_info, "baseUrl\":\"", "\"");              //获取url

                        //new Log(office_file_root_url);

                        return latest_version;
                    }
                    else
                    { return null; }
                }

                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }

            }

            /// <summary>
            /// 获取文件下载列表
            /// </summary>
            /// <returns></returns>
            internal static List<string> GetOfficeFileList()
            {
                try
                {
                    Version version_info = GetOfficeVersion();      //获取版本

                    if (version_info == null || string.IsNullOrEmpty(office_file_root_url))     //下载根地址为空时，视为失败
                    {
                        new Log("     × 最新版本获取失败，请稍后重试！", ConsoleColor.DarkRed);
                        return null;
                    }

                    string ver = version_info.ToString();

                    new Log($"     √ 已完成，最新版：{ver}。", ConsoleColor.DarkGreen);

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
                    new Log($"\n------> 预计下载文件：");
                    foreach (var a in file_list)
                    {
                        new Log($"      > {a}");
                    }
                    Console.Read();
                    */

                    return file_list;
                }

                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
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
            /// Office 本地安装情况（标记）
            /// </summary>
            [Flags]
            internal enum InstallState
            {
                /// <summary>
                /// 已安装了最新版 Office
                /// </summary>
                Correct = 1,

                /// <summary>
                /// 安装了和服务器一样的 Office 架构，但不是最新版
                /// </summary>
                Diff = 2,

                /// <summary>
                /// 系统中存在多个架构版本的 Office
                /// </summary>
                Multi = 4,

                /// <summary>
                /// 本机未安装任何 Office 或 没有使用本软件安装 Office ，通过是否存在 ClickToRun 项判断
                /// </summary>
                None = 8,
            }

            /// <summary>
            /// 判断当前安装的Office版本是否是最新版
            /// （只检查注册表，不包含激活许可证信息）
            /// </summary>
            /// <returns></returns>
            internal static InstallState GetOfficeState()
            {
                var version_info = GetArchiDir();

                InstallState result;

                if (version_info != null && version_info.Count > 0)
                {
                    string office_reg_ver = Register.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "VersionToReport");

                    //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                    string info = Lib_AppUpdate.latest_info;
                    if (string.IsNullOrEmpty(info))
                    {
                        info = Com_WebOS.Visit_WebClient(Lib_AppUpdate.update_json_url);
                    }
                    //获取目标ID
                    string Pop_Office_ID = Com_TextOS.GetCenterText(info, "\"Pop_Office_ID\": \"", "\"");
                    //获取失败时，默认使用 ProPlus2021Volume 版
                    if (string.IsNullOrEmpty(Pop_Office_ID))
                    {
                        Pop_Office_ID = "ProPlus2021Volume";
                    }

                    //获取本机已经安装的ID
                    string Current_Office_ID = Register.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                    if (
                        (office_reg_ver != null && office_reg_ver == OfficeNetVersion.latest_version.ToString())    //版本号一致
                        &&
                        (!string.IsNullOrEmpty(Current_Office_ID) && Current_Office_ID == Pop_Office_ID)            //版本ID一致
                        )      //必须先判断不为null，否则会抛出异常
                    {
                        //已经正确安装最新版
                        result = InstallState.Correct;
                    }
                    else
                    {
                        //版本号和一开始下载的版本号不一致，或者 office_reg_ver 为空，视为版本不同
                        result = InstallState.Diff;
                    }

                    //判断是否只安装了一个版本
                    if (version_info.Count > 1)
                    {
                        //安装了多个版本
                        result |= InstallState.Multi;   //添加标记
                    }
                }
                else
                {
                    //未安装任何
                    result = InstallState.None;
                }

                return result;
            }

            /// <summary>
            /// 获取本机已经安装的所有Office版本和对应的安装目录（字典）
            /// </summary>
            /// <returns></returns>
            internal static Dictionary<OfficeArchi, string> GetArchiDirFull()
            {
                try
                {
                    //注册表根路径
                    string office_reg_root = "SOFTWARE\\Microsoft\\Office";

                    //获取所有版本 office 的安装目录
                    Dictionary<OfficeArchi, string> office_installed_dir = new Dictionary<OfficeArchi, string>
                    {
                        {OfficeArchi.Office_ODT,  Register.GetValue($@"{office_reg_root}\ClickToRun","InstallPath") },                          //2016及其以上版本，通过odt安装的office
                        {OfficeArchi.Office_2013, Register.GetValue($@"{office_reg_root}\15.0\Common\InstallRoot", "Path" ) },                  //2013版
                        {OfficeArchi.Office_2010, Register.GetValue($@"{office_reg_root}\14.0\Common\InstallRoot", "Path") },                   //2010版
                        {OfficeArchi.Office_2007, Register.GetValue($@"{office_reg_root}\12.0\Common\InstallRoot", "Path" ) },                  //2007版              
                    };

                    return office_installed_dir;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }

            /// <summary>
            /// 获取本机已经安装的所有Office版本的安装目录（仅包含有效目录）
            /// </summary>
            /// <returns></returns>
            internal static List<string> GetArchiDir()
            {
                try
                {
                    //先获取完整信息
                    var office_installed_dir_full = GetArchiDirFull();

                    //无目录返回null
                    if (office_installed_dir_full == null)
                    {
                        return null;
                    }

                    //遍历所有目录，验证有效性
                    List<string> office_installed_dir = new List<string>();     //获取所有已知版本的安装路径//将完整信息中的路径作为List存储
                    foreach (var now_dir in office_installed_dir_full)
                    {
                        //非空判断
                        if (!string.IsNullOrEmpty(now_dir.Value))
                        {
                            if (Directory.Exists(now_dir.Value))
                            {
                                office_installed_dir.Add(now_dir.Value);
                            }
                        }
                    }

                    return office_installed_dir;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }

            /// <summary>
            /// 获取目前所有激活、未激活的信息
            /// </summary>
            /// <returns></returns>
            internal static List<string> LicenseInfo()
            {
                try
                {
                    string cmd_switch_cd = $"pushd \"{App.Path.Dir_SDK + @"\Activate"}\"";                  //切换至OSPP文件目录
                    string cmd_installed_info = "cscript ospp.vbs /dstatus";                                //查看激活状态
                    string detect_info = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_installed_info})");     //查看所有版本激活情况

                    //未安装key，直接返回null
                    if (detect_info.Contains("No installed product keys detected"))
                    {
                        return null;
                    }

                    string[] info = detect_info.Split('\n');
                    List<string> key_list = new List<string>();
                    if (info.Length == 0)
                    {
                        return null;
                    }
                    else
                    {
                        foreach (var now_line in info)
                        {
                            if (now_line.Contains("key:"))
                            {
                                string now_key = Com_TextOS.GetCenterText(now_line, "key:", "").Replace(" ", "").TrimEnd('\r');     //需要移除尾部 \r 的标记
                                key_list.Add(now_key);
                            }
                        }
                        return key_list;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }
    }
}
