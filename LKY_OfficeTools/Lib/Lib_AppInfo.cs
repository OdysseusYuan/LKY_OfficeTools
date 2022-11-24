/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Diagnostics;
using System.IO;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 公共信息类库
    /// </summary>
    internal class Lib_AppInfo
    {
        /// <summary>
        /// App 属性类，常量存储
        /// </summary>
        internal class AppAttribute
        {
            /// <summary>
            /// APP 名称
            /// </summary>
            internal const string AppName = "LKY Office Tools";

            /// <summary>
            /// APP 服务的名称（不含空格，避免后续麻烦）
            /// </summary>
            internal static readonly string ServiceName = ServiceDisplayName.Replace(" ", "");

            /// <summary>
            /// APP 服务展示的名称
            /// </summary>
            internal const string ServiceDisplayName = AppName + " Service";

            /// <summary>
            /// APP 版本号
            /// </summary>
            internal const string AppVersion = "1.1.0.21123";

            /// <summary>
            /// 开发者拼音全拼
            /// </summary>
            internal const string Developer = "LiuKaiyuan";
        }

        /// <summary>
        /// App 用于方便开发的相关类库
        /// </summary>
        internal class AppDevelop
        {
            /// <summary>
            /// 全局父级（顶级）命名空间
            /// </summary>
            internal const string NameSpace_Top = "LKY_OfficeTools";
        }

        /// <summary>
        /// App 路径类，动态变量
        /// </summary>
        internal class AppPath
        {
            /// <summary>
            /// 程序运行的目录路径（不是工作目录，是路径的目录）
            /// </summary>
            internal static string ExecuteDir
            {
                get
                {
                    return AppDomain.CurrentDomain.BaseDirectory;
                }
            }

            /// <summary>
            /// 当前 APP 自身所在的完整路径（含自身文件名）
            /// </summary>
            internal static string Executer
            {
                get
                {
                    return new FileInfo(Process.GetCurrentProcess().MainModule.FileName).FullName;
                }
            }

            /// <summary>
            /// APP 文档目录信息
            /// </summary>
            internal class Documents
            {
                /// <summary>
                /// APP 文档根目录。
                /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools
                /// </summary>
                internal static string Documents_Root
                {
                    get
                    {
                        return $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\{AppAttribute.AppName}";
                    }
                }

                /// <summary>
                /// APP 服务运行目录 大类
                /// </summary>
                internal class Services
                {
                    /// <summary>
                    /// 服务目录的根目录。
                    /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools\Services
                    /// </summary>
                    internal static string Services_Root
                    {
                        get
                        {
                            return $"{Documents_Root}\\Services";
                        }
                    }

                    /// <summary>
                    /// 存储服务触发的 被动安装exe的进程信息。
                    /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools\Services\PassiveProcess.info
                    /// </summary>
                    internal static string PassiveProcessInfo
                    {
                        get
                        {
                            return $"{Services_Root}\\PassiveProcess.info";
                        }
                    }
                }

                /// <summary>
                /// APP 日志存储目录。
                /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools\Logs
                /// </summary>
                internal static string Logs
                {
                    get
                    {
                        return $"{Documents_Root}\\Logs";
                    }
                }

                /// <summary>
                /// APP 临时文件夹目录。
                /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools\Temp
                /// </summary>
                internal static string Temp
                {
                    get
                    {
                        return $"{Documents_Root}\\Temp";
                    }
                }

                /// <summary>
                /// SDK 文件子目录
                /// </summary>
                internal class SDKs
                {
                    /// <summary>
                    /// APP SDK文件夹目录。
                    /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools\SDKs
                    /// </summary>
                    internal static string SDKs_Root
                    {
                        get
                        {
                            return $"{Documents_Root}\\SDKs";
                        }
                    }

                    /// <summary>
                    /// Activate 激活文件目录。
                    /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools\Activate
                    /// </summary>
                    internal static string Activate
                    {
                        get
                        {
                            return $"{SDKs_Root}\\Activate";
                        }
                    }

                    /// <summary>
                    /// OSPP 文件路径。
                    /// C:\Users\Odysseus.Yuan\Documents\LKY Office Tools\Activate\OSPP.VBS
                    /// </summary>
                    internal static string Activate_OSPP
                    {
                        get
                        {
                            return $"{Activate}\\OSPP.VBS";
                        }
                    }
                }
            }
        }
    }
}
