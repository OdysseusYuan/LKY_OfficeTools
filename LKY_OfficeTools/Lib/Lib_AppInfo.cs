/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Diagnostics;
using System.IO;
using static LKY_OfficeTools.Lib.Lib_AppLog;

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
            /// APP 名称（简称）
            /// </summary>
            internal const string AppName_Short = "LOT";

            /// <summary>
            /// APP 服务的名称（不含空格，避免后续麻烦）
            /// </summary>
            internal static readonly string ServiceName = ServiceDisplayName.Replace(" ", "");

            /// <summary>
            /// APP 服务展示的名称
            /// </summary>
            internal const string ServiceDisplayName = AppName + " Service";

            /// <summary>
            /// APP 文件名
            /// </summary>
            internal const string AppFilename = "LKY_OfficeTools.exe";

            /// <summary>
            /// APP 版本号
            /// </summary>
            internal const string AppVersion = "1.1.1.1209";

            /// <summary>
            /// 开发者拼音全拼
            /// </summary>
            internal const string Developer = "LiuKaiyuan";
        }

        /// <summary>
        /// Json 信息
        /// </summary>
        internal class AppJson
        {
            /// <summary>
            /// AppJson.Info 的私有变量。
            /// 用于寄存其公有变量的首次赋值信息，以免每次访问公有变量时，都要重新读取网页，达到节省资源目的。
            /// </summary>
            private static string AppJsonInfo = null;

            /// <summary>
            /// 获取 LKY_OfficeTools_AppInfo.json 的信息
            /// </summary>
            internal static string Info
            {
                get
                {
                    try
                    {
                        //内部变量非空时，直接返回其值
                        if (!string.IsNullOrEmpty(AppJsonInfo))
                        {
                            return AppJsonInfo;
                        }

#if (!DEBUG)
                        //release模式地址
                        string json_url = $"https://gitee.com/OdysseusYuan/LKY_OfficeTools/releases/download/AppInfo/LKY_OfficeTools_AppInfo.json";
#else
                        //debug模式地址
                        string json_url = $"https://gitee.com/OdysseusYuan/LOT_OnlyTest/releases/download/AppInfo_Test/test.json";
#endif

                        //获取 Json 信息
                        string result = Com_WebOS.Visit_WebClient(json_url);

                        AppJsonInfo = result;
                        return AppJsonInfo;
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                        //new Log($"     × 获取 AppJson 信息失败！", ConsoleColor.DarkRed);
                        new Log($"获取 AppJson 信息失败！");
                        return null;
                    }
                }
            }
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
            /// 程序运行的目录路径（不是工作目录，是路径的目录）。
            /// 当程序自身被 move 到别的地方后，该值不会发生改变
            /// </summary>
            internal static string ExecuteDir
            {
                get
                {
                    //使用 BaseDirectory 获取路径时，结尾会多一个 \ 符号，为了替换之，先在结果处增加一个斜杠，变成 \\ 结尾，然后替换掉这个 双斜杠
                    return (AppDomain.CurrentDomain.BaseDirectory + @"\").Replace(@"\\", @"");
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
            /// APP 文档目录信息。
            /// 系统级服务（LocalSystem）模式下，无法获取“我的文档”目录，故换用新目录 ProgramData
            /// </summary>
            internal class Documents
            {
                /// <summary>
                /// APP 文档根目录。
                /// C:\ProgramData\LKY Office Tools
                /// </summary>
                internal static string Documents_Root
                {
                    get
                    {
                        return $"{Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)}\\{AppAttribute.AppName}";     //ProgramData
                        //return $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\{AppAttribute.AppName}";            //我的文档
                    }
                }

                /// <summary>
                /// APP Update 目录大类
                /// </summary>
                internal class Update
                {
                    /// <summary>
                    /// 升级目录的根目录。
                    /// C:\ProgramData\LKY Office Tools\Update
                    /// </summary>
                    internal static string Update_Root
                    {
                        get
                        {
                            return $"{Documents_Root}\\Update";
                        }
                    }

                    /// <summary>
                    /// 升级目录的回收站目录。
                    /// C:\ProgramData\LKY Office Tools\Update\Trash
                    /// </summary>
                    internal static string UpdateTrash
                    {
                        get
                        {
                            return $"{Update_Root}\\Trash";
                        }
                    }
                }

                /// <summary>
                /// APP 服务运行目录 大类
                /// </summary>
                internal class Services
                {
                    /// <summary>
                    /// 服务目录的根目录。
                    /// C:\ProgramData\LKY Office Tools\Services
                    /// </summary>
                    internal static string Services_Root
                    {
                        get
                        {
                            return $"{Documents_Root}\\Services";
                        }
                    }

                    /// <summary>
                    /// 服务目录的回收站目录，用于存放丢弃的文件。
                    /// C:\ProgramData\LKY Office Tools\Services\Trash
                    /// </summary>
                    internal static string ServicesTrash
                    {
                        get
                        {
                            return $"{Services_Root}\\Trash";
                        }
                    }

                    /// <summary>
                    /// 用于存放服务模式运行的文件所在目录。
                    /// 此举的目的，是防止用户手动将当前文件删除，导致服务无法启动。
                    /// C:\ProgramData\LKY Office Tools\Services\Autorun
                    /// </summary>
                    internal static string ServiceAutorun
                    {
                        get
                        {
                            return $"{Services_Root}\\Autorun";
                        }
                    }

                    /// <summary>
                    /// 用于服务模式运行的 exe 路径，文件与手动运行的一模一样，只是位置不同。
                    /// 此举的目的，是防止用户手动将当前文件删除，导致服务无法启动。
                    /// C:\ProgramData\LKY Office Tools\Services\Autorun\LKY_OfficeTools.exe
                    /// </summary>
                    internal static string ServiceAutorun_Exe
                    {
                        get
                        {
                            return $"{ServiceAutorun}\\{AppAttribute.AppFilename}";
                        }
                    }

                    /// <summary>
                    /// 存储服务触发的 被动安装exe的进程信息。
                    /// C:\ProgramData\LKY Office Tools\Services\PassiveProcess.info
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
                /// C:\ProgramData\LKY Office Tools\Logs
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
                /// C:\ProgramData\LKY Office Tools\Temp
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
                    /// C:\ProgramData\LKY Office Tools\SDKs
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
                    /// C:\ProgramData\LKY Office Tools\Activate
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
                    /// C:\ProgramData\LKY Office Tools\Activate\OSPP.VBS
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
