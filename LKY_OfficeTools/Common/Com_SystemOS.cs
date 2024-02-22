/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 LiuKaiyuan. All rights reserved.
 *      
 *      FileName : Com_SystemOS.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    internal class Com_SystemOS
    {
        internal class OSVersion
        {
            internal enum OSType
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

            internal static readonly IDictionary<string, string> WinPublishType = new Dictionary<string, string>
                {
                    //win10版本号
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

                    //win11版本号
                    { "22000", "21H2" },
                    { "22621", "22H2" },

                    //{ "1111111111111111", "LTSB" },
                    //{ "1111111111111111", "LTSC" },
                    //{ "1111111111111111", "ARM" },
                };

            internal static OSType GetPublishType()
            {
                try
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
                        //检查注册表，因为win10和11的主版本号都为10，只能用buildID来判断了
                        string curr_ver = Register.Read.ValueBySystem(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild");

                        if (!string.IsNullOrEmpty(curr_ver) && int.Parse(curr_ver) < 22000)       //Win11目前内部版本号
                        {
                            return OSType.Win10;
                        }
                        else
                        {
                            return OSType.Win11;
                        }
                    }
                    else
                    {
                        return OSType.UnKnow;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return OSType.UnKnow;
                }
            }

            internal static string GetBuildNumber(bool isCoreVersion = true)
            {
                try
                {
                    //检查注册表
                    string curr_mode = Register.Read.ValueBySystem(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild");

                    //为空返回未知
                    if (string.IsNullOrEmpty(curr_mode))
                    {
                        return "unknow";
                    }

                    //判断返回的内容
                    if (isCoreVersion)      //返回内部版本号
                    {
                        return curr_mode;
                    }
                    else                    //返回发行版本
                    {
                        return WinPublishType[curr_mode];
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return "error!";
                }
            }
        }

        internal class Register
        {
            internal class Read
            {
                internal static string Value(RegistryHive reg_root, RegistryView reg_view, string path, string key)
                {
                    try
                    {
                        RegistryKey HK_Root = RegistryKey.OpenBaseKey(reg_root, reg_view);

                        RegistryKey path_reg = HK_Root.OpenSubKey(path);    //先获取路径

                        if (path_reg == null)
                        {
                            //找不到注册表路径
                            return null;
                        }
                        else
                        {
                            object value = path_reg.GetValue(key);
                            if (value != null)      //必须先判断不为null，否则会抛出异常
                            {
                                //一切正常
                                return value.ToString();
                            }
                            else
                            {
                                //Key不存在或值为空
                                return null;
                            }
                        }
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                        return null;
                    }
                }

                internal static string ValueBySystem(RegistryHive reg_root, string path, string key)
                {
                    try
                    {
                        if (Environment.Is64BitOperatingSystem)
                        {
                            //x64系统，访问x64注册表
                            return Value(reg_root, RegistryView.Registry64, path, key);
                        }
                        else
                        {
                            //x32系统，访问x32注册表
                            return Value(reg_root, RegistryView.Registry32, path, key);
                        }
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                        return null;
                    }
                }

                internal static List<string> AllValues(RegistryHive reg_root, string path, string key)
                {
                    try
                    {
                        List<string> result = new List<string>();

                        //获取x32的结构值
                        string value_x32 = Value(reg_root, RegistryView.Registry32, path, key);
                        if (!string.IsNullOrWhiteSpace(value_x32))
                        {
                            result.Add(value_x32);
                        }

                        //获取x64的结构值（仅在当前计算机为 x64 系统时，才获取）
                        if (Environment.Is64BitOperatingSystem)
                        {
                            string value_x64 = Value(reg_root, RegistryView.Registry64, path, key);
                            if (!string.IsNullOrWhiteSpace(value_x64))
                            {
                                result.Add(value_x64);
                            }
                        }

                        return result;
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                        return null;
                    }
                }
            }

            internal static bool ExistItem(RegistryHive reg_root, RegistryView reg_view, string item_path)
            {
                try
                {
                    RegistryKey reg = RegistryKey.OpenBaseKey(reg_root, reg_view);
                    var result = reg.OpenSubKey(item_path);

                    if (result != null)
                    {
                        return true;
                    }

                    result.Close();

                    return false;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            internal static bool DeleteItem(RegistryHive reg_root, RegistryView reg_view, string path, string item)
            {
                try
                {
                    RegistryKey HK_Root = RegistryKey.OpenBaseKey(reg_root, reg_view);
                    RegistryKey path_reg = HK_Root.OpenSubKey(path, true);              //先获取路径，启动可写模式

                    //找不到注册表路径，默认已删除，返回true
                    if (path_reg == null)
                    {
                        return true;
                    }

                    path_reg.DeleteSubKeyTree(item, false);

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            internal static bool ExportReg(string reg_path, string save_to)
            {
                try
                {
                    //先删除已经存在的文件
                    if (File.Exists(save_to))
                    {
                        File.Delete(save_to);
                    }

                    //开始导出
                    Directory.CreateDirectory(new FileInfo(save_to).DirectoryName);         //先创建文件所在目录
                    Process.Start("regedit", $" /E \"{save_to}\" \"{reg_path}\"").WaitForExit();

                    if (File.Exists(save_to))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }
        }

        internal class SoftWare
        {
            public static List<string> InstalledList()
            {
                try
                {
                    //从注册表中获取控制面板“卸载程序”中的程序和功能列表
                    RegistryKey HK_Root_x32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);    //打开x32系统键
                    RegistryKey HK_Root_x64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);    //打开x64系统键

                    string soft_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall";
                    RegistryKey software_x32_list = HK_Root_x32.OpenSubKey(soft_path);              //x32软件列表
                    RegistryKey software_x64_list = null;                                           //x64软件列表。x32系统该值将始终为null

                    //获取x64软件列表（仅在用户系统是x64的情况下获取）
                    if (Environment.Is64BitOperatingSystem)
                    {
                        software_x64_list = HK_Root_x64.OpenSubKey(soft_path);
                    }

                    //整合键位
                    List<RegistryKey> software_key = new List<RegistryKey>();
                    if (software_x32_list != null && software_x32_list.SubKeyCount > 0)
                    {
                        software_key.Add(software_x32_list);
                    }
                    if (software_x64_list != null && software_x64_list.SubKeyCount > 0)
                    {
                        software_key.Add(software_x64_list);
                    }

                    //开始获取
                    if (software_key != null && software_key.Count > 0)
                    {
                        List<string> software_info = new List<string>();

                        foreach (var now_bit in software_key)                               //遍历2个系统位数的注册表
                        {
                            foreach (string now_subkeyname in now_bit.GetSubKeyNames())     //遍历每个位数下面，对应的软件列表
                            {
                                //打开对应的软件名称
                                RegistryKey SubKey = now_bit.OpenSubKey(now_subkeyname);
                                if (SubKey != null)
                                {
                                    string DisplayName = SubKey.GetValue("DisplayName", "NONE").ToString();

                                    //过滤条件
                                    if (DisplayName != "NONE" && !DisplayName.Contains("vs") && !DisplayName.Contains("Visual C++") &&
                                        !DisplayName.Contains(".NET"))
                                    {
                                        software_info.Add(DisplayName);
                                    }
                                }
                            }
                        }

                        //元素去重
                        if (software_info != null && software_info.Count > 0)
                        {
                            return software_info.Distinct().ToList();
                        }
                    }

                    return null;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }

        internal class Screen
        {
            internal static bool CaptureToSave(string save_to, ImageFormat file_type = null)
            {
                try
                {
                    //初始化屏幕尺寸
                    int screenLeft = SystemInformation.VirtualScreen.Left;
                    int screenTop = SystemInformation.VirtualScreen.Top;
                    int screenWidth = SystemInformation.VirtualScreen.Width;
                    int screenHeight = SystemInformation.VirtualScreen.Height;

                    //接收截图
                    using (Bitmap bmp = new Bitmap(screenWidth, screenHeight))
                    {
                        //抓取
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.CopyFromScreen(screenLeft, screenTop, 0, 0, bmp.Size);
                        }

                        //判断保存格式
                        if (file_type == null)
                        {
                            file_type = ImageFormat.Jpeg;
                        }

                        //创建日志目录
                        Directory.CreateDirectory(new FileInfo(save_to).DirectoryName);

                        bmp.Save(save_to, file_type);
                    }

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }
        }
    }
}