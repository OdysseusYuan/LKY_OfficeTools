/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_SystemOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 获取用户电脑信息类库
    /// </summary>
    internal class Com_SystemOS
    {
        /// <summary>
        /// 系统级别类库
        /// </summary>
        internal class OSVersion
        {
            /// <summary>
            /// 操作系统类别枚举
            /// </summary>
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

            /// <summary>
            /// Windows10~11版本号与发行号对应字典
            /// 官方字典：https://docs.microsoft.com/zh-cn/windows/release-health/release-information
            /// </summary>
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

            /// <summary>
            /// 判断用户电脑Windows操作系统版本，如：Win7、Win10等
            /// </summary>
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
                        string curr_ver = Registry.GetValue(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild");

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
                catch
                {
                    return OSType.UnKnow;
                }
            }

            /// <summary>
            /// 获取当前Win BuildNumber，isCoreVersion默认为真，返回内部版本号，如：19043，
            /// 若为假，则返回发行版本，如：21H1版本。
            /// </summary>
            internal static string GetBuildNumber(bool isCoreVersion = true)
            {
                try
                {
                    //检查注册表
                    string curr_mode = Registry.GetValue(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild");

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
                catch
                {
                    return "error!";
                }
            }            
        }

        /// <summary>
        /// 注册表操作类库
        /// </summary>
        internal class Registry
        {
            /// <summary>
            /// 获取指定路径下的注册表键值。
            /// 默认注册表根部项为 HKLM
            /// </summary>
            /// <returns></returns>
            internal static string GetValue(string path, string key, RegistryHive root = RegistryHive.LocalMachine)
            {
                try
                {
                    RegistryKey HK_Root = RegistryKey.OpenBaseKey(root,
                        Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);      //判断操作系统版本（64位\32位）打开注册表项，不然 x86编译的本程序 读取 x64的程序会出现无法读取 已经存在于注册表 中的数据

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
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 软件操作类库
        /// </summary>
        internal class SoftWare
        {
            /// <summary>
            /// 获取已安装软件列表
            /// </summary>
            /// <returns></returns>
            public static List<string> GetList()
            {
                try
                {
                    //从注册表中获取控制面板“卸载程序”中的程序和功能列表
                    RegistryKey Key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                        "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
                    if (Key != null) //如果系统禁止访问则返回null
                    {
                        List<string> software_info = new List<string>();

                        foreach (string SubKeyName in Key.GetSubKeyNames())
                        {
                            //打开对应的软件名称
                            RegistryKey SubKey = Key.OpenSubKey(SubKeyName);
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

                        return software_info;
                    }
                    return null;
                }
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 屏幕相关类库
        /// </summary>
        internal class Screen
        {
            /// <summary>
            /// 抓取屏幕并保存至文件
            /// </summary>
            /// <param name="save_to"></param>
            /// <param name="file_type"></param>
            /// <returns></returns>
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
                            file_type = ImageFormat.Png;
                        }

                        bmp.Save(save_to, file_type);
                    }

                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
    }
}
