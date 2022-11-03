/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_SystemOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using Microsoft.Win32;
using System;
using System.Collections.Generic;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 获取用户电脑信息类库
    /// </summary>
    public class Com_SystemOS
    {
        /// <summary>
        /// 系统级别类库
        /// </summary>
        public class OS
        {
            /// <summary>
            /// 操作系统类别枚举
            /// </summary>
            public enum OSType
            {
                LowVersion,
                WinXP,
                WinVista,
                Win7,       //支持.NET 4.8
                Win8_1,     //支持.NET 4.8
                Win10,      //1607开始，支持.NET 4.8
                Win11,
                UnKnow
            }

            /// <summary>
            /// Windows10版本号与发行号对应字典
            /// 官方字典：https://docs.microsoft.com/zh-cn/windows/release-health/release-information
            /// </summary>
            public static readonly IDictionary<string, string> Win10Type = new Dictionary<string, string>
                {
                    { "10240", "1507" },
                    { "10586", "1511" },
                    { "14393", "1607" },        //.NET 4.8从此版本开始支持
                    { "15063", "1703" },
                    { "16299", "1709" },
                    { "17134", "1803" },
                    { "17763", "1809" },
                    { "18362", "1903" },
                    { "18363", "1909" },
                    { "19041", "2004" },
                    { "19042", "20H2" },
                    { "19043", "21H1" },
                    { "19044", "21H2" },
                    { "19045", "22H2" },
                    //{ "1111111111111111", "LTSB" },
                    //{ "1111111111111111", "LTSC" },
                    //{ "1111111111111111", "ARM" },
                };

            /// <summary>
            /// 获取当前Win10子版本，isCoreVersion默认为真，返回内部版本号，如：19043，
            /// 若为假，则返回发行版本，如：21H1版本。
            /// </summary>
            public static string Win10Version(bool isCoreVersion = true)
            {
                //检查注册表
                RegistryKey HKLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                        Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);      //判断操作系统版本（64位\32位）打开注册表项，不然 x86编译的本程序 读取 x64的程序会出现无法读取 已经存在于注册表 中的数据
                string reg_path = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";
                RegistryKey office_reg = HKLM.OpenSubKey(reg_path);
                string curr_mode = office_reg.GetValue("CurrentBuild").ToString();

                if (isCoreVersion)      //返回内部版本号
                {
                    return curr_mode;
                }
                else if (!string.IsNullOrEmpty(curr_mode))      //返回发行版本
                {
                    return Win10Type[curr_mode];
                }
                else { return null; }
            }

            /// <summary>
            /// Windows11版本号与发行号对应字典
            /// 官方字典：https://learn.microsoft.com/zh-cn/windows/release-health/windows11-release-information
            /// </summary>
            public static readonly IDictionary<string, string> Win11Type = new Dictionary<string, string>
                {
                    { "22000", "21H2" },
                    { "22621", "22H2" },
                    //{ "1111111111111111", "LTSB" },
                    //{ "1111111111111111", "LTSC" },
                    //{ "1111111111111111", "ARM" },
                };

            /// <summary>
            /// 获取当前Win11子版本，isCoreVersion默认为真，返回内部版本号，如：19043，
            /// 若为假，则返回发行版本，如：21H1版本。
            /// </summary>
            public static string Win11Version(bool isCoreVersion = true)
            {
                //检查注册表
                RegistryKey HKLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                        Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);      //判断操作系统版本（64位\32位）打开注册表项，不然 x86编译的本程序 读取 x64的程序会出现无法读取 已经存在于注册表 中的数据
                string reg_path = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";
                RegistryKey office_reg = HKLM.OpenSubKey(reg_path);
                string curr_mode = office_reg.GetValue("CurrentBuild").ToString();

                if (isCoreVersion)      //返回内部版本号
                {
                    return curr_mode;
                }
                else if (!string.IsNullOrEmpty(curr_mode))      //返回发行版本
                {
                    return Win11Type[curr_mode];
                }
                else { return null; }
            }

            /// <summary>
            /// 判断用户电脑Windows操作系统版本，如：Win7、Win10等
            /// </summary>
            public static OSType WindowsVersion()
            {
                Version ver = Environment.OSVersion.Version;

                if (ver.Major < 5)
                {
                    return OSType.LowVersion;
                }
                else if (ver.Major == 5 && ver.Minor == 1)
                {
                    return OSType.WinXP;
                }
                else if (ver.Major == 6 && ver.Minor == 0)
                {
                    return OSType.WinVista;
                }
                else if (ver.Major == 6 && ver.Minor == 1)
                {
                    return OSType.Win7;
                }
                else if (ver.Major == 6 && ver.Minor == 2)
                {
                    return OSType.Win8_1;
                }
                else if (ver.Major == 10 && ver.Minor == 0)     //正确获取win10版本号，需要在exe里面加入app.manifest
                {
                    return OSType.Win10;
                }
                else
                {
                    //检查注册表
                    RegistryKey HKLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                            Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);      //判断操作系统版本（64位\32位）打开注册表项，不然 x86编译的本程序 读取 x64的程序会出现无法读取 已经存在于注册表 中的数据
                    string reg_path = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";
                    RegistryKey office_reg = HKLM.OpenSubKey(reg_path);
                    string curr_ver = office_reg.GetValue("CurrentBuild").ToString();

                    if (!string.IsNullOrEmpty(curr_ver) && int.Parse(curr_ver) >= 22000)       //Win11目前内部版本号
                    {
                        return OSType.Win11;
                    }
                    else
                    {
                        return OSType.UnKnow;
                    }
                }
            }
        }
    }
}
