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
using System.IO;
using System.Linq;
using System.Threading;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.AppPath;
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
            /// Office 2003
            /// </summary>
            Office_2003 = 2003,

            /// <summary>
            /// Office 12.0 x32
            /// </summary>
            Office_2007_x32 = 2007 * 32,

            /// <summary>
            /// Office 12.0 x64
            /// </summary>
            Office_2007_x64 = 2007 * 64,

            /// <summary>
            /// Office 14.0 x32
            /// </summary>
            Office_2010_x32 = 2010 * 32,

            /// <summary>
            /// Office 14.0 x64
            /// </summary>
            Office_2010_x64 = 2010 * 64,

            /// <summary>
            /// Office 15.0 x32
            /// </summary>
            Office_2013_x32 = 2013 * 32,

            /// <summary>
            /// Office 15.0 x64
            /// </summary>
            Office_2013_x64 = 2013 * 64,

            /// <summary>
            /// Office 16.0及其以上版本  x32
            /// </summary>
            Office_ODT_x32 = 2016 * 32,

            /// <summary>
            /// Office 16.0及其以上版本 x64
            /// </summary>
            Office_ODT_x64 = 2016 * 64,
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

            /// <summary>
            /// 要下载的 Office 文件的根网址
            /// </summary>
            internal static string office_file_root_url;

            /// <summary>
            /// 获取最新版本的 Office 函数
            /// </summary>
            /// <returns></returns>
            internal static Version GetOfficeVersion()
            {
                try
                {
                    new Log("\n------> 正在获取 最新可用 Office 版本 ...", ConsoleColor.DarkCyan);

                    //获取频道信息，得到可用 Url 列表
                    string channel_info = Com_TextOS.GetCenterText(AppJson.Info, "\"OfficeChannels_Url\": \"", "\"");
                    if (string.IsNullOrEmpty(channel_info))
                    {
                        throw new Exception("OfficeChannels_Url 信息获取失败！");
                    }
                    List<string> channel_list = channel_info.Split(';').ToList();

                    //遍历获取有效地址，并下载channel信息
                    string office_info = null;
                    int try_times = 1;                                                          //尝试获取次数，初始值为1
                    foreach (var now_url in channel_list)
                    {
                        office_info = Com_WebOS.Visit_WebClient(now_url.Replace(" ", ""));      //替换网址空格，并获得当前网址的访问信息
                        if (!string.IsNullOrEmpty(office_info))
                        {
                            //info有值，开始截取字段

                            //获取版本信息
                            string latest_info = Com_TextOS.GetCenterText(office_info, "\"PerpetualVL2021\",", "name");                     //获取 2021 LTSC
                            if (!string.IsNullOrEmpty(latest_info))
                            {
                                //获取版本号
                                latest_version = new Version(Com_TextOS.GetCenterText(latest_info, "latestUpdateVersion\":\"", "\""));      //官方Json取值，格式和自己的不一样，千万注意
                                
                                //获取office下载地址
                                office_file_root_url = Com_TextOS.GetCenterText(latest_info, "baseUrl\":\"", "\"");                         //官方Json取值，格式和自己的不一样，千万注意
                                
                                //版本信息不为空，下载地址不为空时，返回
                                if ((latest_version != null) && (!string.IsNullOrEmpty(office_file_root_url)))
                                {
                                    return latest_version;
                                }
                            }
                        }

                        //访问当前网址返回数据为空，或者无法截取获得有效字段时，遍历。
                        if (try_times < channel_list.Count)
                        {
                            new Log($"     >> 尝试使用 {++try_times} 号服务器获取中 ...", ConsoleColor.DarkYellow);
                            continue;
                        }
                    }

                    //全部遍历完，依旧没返回，就说明全部失败，返回null。
                    return null;
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

                    new Log($"     √ 已获得 Office v{ver} 正版信息。", ConsoleColor.DarkGreen);

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
                /// 安装了 Office，但版本与预期版本不同
                /// </summary>
                Diff = 2,

                /// <summary>
                /// 系统中存在多个版本的 Office
                /// </summary>
                Multi = 4,

                /// <summary>
                /// 本机未安装任何 Office
                /// </summary>
                None = 8,
            }

            /// <summary>
            /// 判断当前电脑 Office 版本的安装情况。
            /// （只检查注册表，不包含激活许可证信息）
            /// </summary>
            /// <returns></returns>
            internal static InstallState GetOfficeState()
            {
                var version_info = GetArchiDir();

                if (version_info != null && version_info.Count > 0)     //系统已安装至少1个Office版本
                {
                    //判断是否只安装了一个版本
                    if (version_info.Count > 1)
                    {
                        //安装了多个版本
                        return InstallState.Multi;
                    }
                    else
                    {
                        //系统只有一个版本

                        //获取ODT双位数的版本信息列表（x32、x64）
                        var office_reg_ver_list = Register.Read.AllValues(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "VersionToReport");
                        if (office_reg_ver_list != null && office_reg_ver_list.Count > 0)
                        {
                            //遍历查询安装目录
                            if (office_reg_ver_list.Count == 1)
                            {
                                //只安装了1个ODT版本
                                foreach (var now_dir in office_reg_ver_list)
                                {
                                    //获取目标ID
                                    string Pop_Office_ID = Com_TextOS.GetCenterText(AppJson.Info, "\"Pop_Office_ID\": \"", "\"");
                                    //获取失败时，默认使用 ProPlus2021Volume 版
                                    if (string.IsNullOrEmpty(Pop_Office_ID))
                                    {
                                        Pop_Office_ID = "ProPlus2021Volume";
                                    }

                                    //获取本机已经安装的ID（如果x64系统，安装了x32的 Office 在 WOW6432Node 目录，该值将为null）
                                    string Current_Office_ID = Register.Read.ValueBySystem(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                                    if (
                                        (now_dir != null && now_dir == OfficeNetVersion.latest_version.ToString())          //版本号一致
                                        &&
                                        (!string.IsNullOrEmpty(Current_Office_ID) && Current_Office_ID == Pop_Office_ID)    //产品ID一致
                                        )      //必须先判断不为null，否则会抛出异常
                                    {
                                        //x64系统还得校验下 Office 注册表的 平台信息 是否一致
                                        if (Environment.Is64BitOperatingSystem)
                                        {
                                            string platform = Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry64, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "Platform");
                                            if (string.IsNullOrWhiteSpace(platform) || platform == "x86")
                                            {
                                                return InstallState.Diff;       //虽然出现在了与x64系统注册表匹配的路径，但是Office平台版本并非x64
                                            }
                                            else
                                            {
                                                //版本号一致、产品ID一致、注册表显示的位数一致
                                                return InstallState.Correct;    //已经正确安装最新版
                                            }
                                        }
                                        else
                                        {
                                            //x32系统，直接满足了
                                            return InstallState.Correct;    //已经正确安装最新版
                                        }
                                    }
                                }

                                //ODT 只安装了1个版本，但版本号和预期版本号不一致，或者 now_dir 为空，视为版本不同
                                return InstallState.Diff;
                            }
                            else
                            {
                                //存在多个ODT版本（如：version信息为1，但odt的path信息有2个），不管版本是否相同，都设置 Multi 标记
                                return InstallState.Multi;
                            }
                        }
                        else
                        {
                            //安装了1个Office版本，但ODT版本安装数量为0，则版本不同
                            return InstallState.Diff;
                        }
                    }
                }
                else
                {
                    //未安装任何
                    return InstallState.None;
                }
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

                    //先获取 x32 的，并填充 null 到 x64
                    Dictionary<OfficeArchi, string> office_installed_dir = new Dictionary<OfficeArchi, string>
                    {
                        //2016及其以上版本，通过odt安装的office
                        {OfficeArchi.Office_ODT_x64, null},
                        {OfficeArchi.Office_ODT_x32, Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry32, $@"{office_reg_root}\ClickToRun", "InstallPath") },

                        //2013版
                        {OfficeArchi.Office_2013_x64, null},
                        {OfficeArchi.Office_2013_x32, Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry32, $@"{office_reg_root}\15.0\Common\InstallRoot", "Path" ) },

                        //2010版
                        {OfficeArchi.Office_2010_x64, null},
                        {OfficeArchi.Office_2010_x32, Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry32, $@"{office_reg_root}\14.0\Common\InstallRoot", "Path") },

                        //2007版
                        {OfficeArchi.Office_2007_x64, null},
                        {OfficeArchi.Office_2007_x32, Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry32, $@"{office_reg_root}\12.0\Common\InstallRoot", "Path") },

                        //2003版
                        {OfficeArchi.Office_2003, Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry32, $@"{office_reg_root}\11.0\Common\InstallRoot", "Path") },
                    };

                    //只有x64系统才获取x64注册表值
                    if (Environment.Is64BitOperatingSystem)
                    {
                        //2016及其以上版本
                        office_installed_dir[OfficeArchi.Office_ODT_x64] = Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry64, $@"{office_reg_root}\ClickToRun", "InstallPath");

                        //2013版
                        office_installed_dir[OfficeArchi.Office_2013_x64] = Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry64, $@"{office_reg_root}\15.0\Common\InstallRoot", "Path");

                        //2010版
                        office_installed_dir[OfficeArchi.Office_2010_x64] = Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry64, $@"{office_reg_root}\14.0\Common\InstallRoot", "Path");

                        //2007版
                        office_installed_dir[OfficeArchi.Office_2007_x64] = Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry64, $@"{office_reg_root}\12.0\Common\InstallRoot", "Path");
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
            /// 获取本机已经安装的所有Office版本的安装目录（仅包含有效目录）
            /// </summary>
            /// <returns></returns>
            internal static List<string> GetArchiDir()
            {
                try
                {
                    //先获取完整信息
                    var office_installed_dir_full = GetArchiDirFull();

                    //无目录返回 0 个
                    if (office_installed_dir_full == null)
                    {
                        return new List<string>();
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
                    return new List<string>();
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
                    string cmd_switch_cd = $"pushd \"{Documents.SDKs.Activate}\"";                  //切换至OSPP文件目录
                    string cmd_installed_info = "cscript ospp.vbs /dstatus";                                //查看激活状态
                    string detect_info = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_installed_info})");     //查看所有版本激活情况

                    //未安装key，直接返回 0 个
                    if (detect_info.Contains("No installed product keys detected"))
                    {
                        return new List<string>();
                    }

                    string[] info = detect_info.Split('\n');
                    List<string> key_list = new List<string>();
                    if (info.Length == 0)
                    {
                        return new List<string>();
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
                    return new List<string>();
                }
            }
        }
    }
}
