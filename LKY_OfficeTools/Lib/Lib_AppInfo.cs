/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppReport;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 公共信息类库
    /// </summary>
    internal class Lib_AppInfo
    {
        /// <summary>
        /// APP 相关的类库
        /// </summary>
        internal class App
        {
            /// <summary>
            /// App 路径类
            /// </summary>
            internal class AppPath
            {
                /// <summary>
                /// 程序运行的目录路径（不是工作目录，是路径的目录）
                /// </summary>
                internal static string Execute
                {
                    get
                    {
                        return AppDomain.CurrentDomain.BaseDirectory;
                    }
                }

                /// <summary>
                /// 当前 APP 自身所在的完整路径（含自身文件名）
                /// </summary>
                internal static string Execute_File
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
                    /// APP 文档根目录
                    /// </summary>
                    internal static string Root
                    {
                        get
                        {
                            return $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\LKY Office Tools";
                        }
                    }

                    /// <summary>
                    /// APP 服务自动运行目录
                    /// </summary>
                    internal static string AutoRun
                    {
                        get
                        {
                            return $"{Root}\\AutoRun";
                        }
                    }

                    /// <summary>
                    /// APP 日志存储目录
                    /// </summary>
                    internal static string Log
                    {
                        get
                        {
                            return $"{Root}\\Logs";
                        }
                    }

                    /// <summary>
                    /// APP 临时文件夹目录
                    /// </summary>
                    internal static string Temp
                    {
                        get
                        {
                            return $"{Root}\\Temp";
                        }
                    }

                    /// <summary>
                    /// SDK 文件子目录
                    /// </summary>
                    internal class SDK
                    {
                        /// <summary>
                        /// APP SDK文件夹目录
                        /// </summary>
                        internal static string Root
                        {
                            get
                            {
                                return $"{Documents.Root}\\SDKs";
                            }
                        }

                        /// <summary>
                        /// Activate 激活文件目录
                        /// </summary>
                        internal static string Activate
                        {
                            get
                            {
                                return $"{Root}\\Activate";
                            }
                        }

                        /// <summary>
                        /// OSPP 文件路径
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

            /// <summary>
            /// 版权类
            /// </summary>
            internal class Copyright
            {
                /// <summary>
                /// 开发者拼音全拼
                /// </summary>
                internal const string Developer = "LiuKaiyuan";
            }

            /// <summary>
            /// 用于方便开发的相关类库
            /// </summary>
            internal class Develop
            {
                /// <summary>
                /// 全局父级（顶级）命名空间
                /// </summary>
                internal const string NameSpace_Top = "LKY_OfficeTools";
            }

            /// <summary>
            /// APP 状态类库
            /// </summary>
            internal class State
            {
                /// <summary>
                /// 当前 APP 运行模式
                /// </summary>
                internal enum RunMode
                {
                    /// <summary>
                    /// 手动模式运行
                    /// </summary>
                    Manual,

                    /// <summary>
                    /// 服务模式运行
                    /// </summary>
                    Service
                }

                /// <summary>
                /// APP 当前运行模式（默认为手动模式）
                /// </summary>
                internal static RunMode Current_RunMode = RunMode.Manual;

                /// <summary>
                /// APP运行状态类型
                /// </summary>
                internal enum ProcessStage
                {
                    /// <summary>
                    /// 等待用户输入（程序未完成情况下）
                    /// </summary>
                    WaitInput = 1,

                    /// <summary>
                    /// 程序正在初始化阶段
                    /// </summary>
                    Starting = 2,

                    /// <summary>
                    /// 系统运行中
                    /// </summary>
                    Process = 4,

                    /// <summary>
                    /// 强制结束，中断运行
                    /// </summary>
                    Interrupt = 8,

                    /// <summary>
                    /// 已结束，且运行成功。即使存在等待用户输入，但程序自身已经完成
                    /// </summary>
                    Finish_Success = 16,

                    /// <summary>
                    /// 已结束，运行有严重错误。导致失败。即使存在等待用户输入，但程序自身已经完成
                    /// </summary>
                    Finish_Fail = 32,
                }

                /// <summary>
                /// APP 当前状态（初始值为 Process）
                /// </summary>
                internal static ProcessStage Current_StageType = ProcessStage.Process;

                /// <summary>
                /// 关闭的方式
                /// </summary>
                internal class Close
                {
                    internal delegate bool ControlCtrlDelegate(int CtrlType);

                    [DllImport("kernel32.dll")]
                    internal static extern bool SetConsoleCtrlHandler(ControlCtrlDelegate HandlerRoutine, bool Add);

                    internal static ControlCtrlDelegate newDelegate = new ControlCtrlDelegate(HandlerRoutine);

                    /// <summary>
                    /// 提示用户非正常关闭
                    /// </summary>
                    /// <param name="CtrlType"></param>
                    /// <returns></returns>
                    internal static bool HandlerRoutine(int CtrlType)
                    {
                        //只要程序不是已完成（无论成功与否），手动关闭，就会显示文字并打点
                        if (Current_StageType != ProcessStage.Finish_Fail && Current_StageType != ProcessStage.Finish_Success)
                        {
                            new Log($"\n     × 正在尝试 取消部署，请稍候 ...", ConsoleColor.Red);
                            Console.ForegroundColor = ConsoleColor.Gray;     //重置颜色，如果第一次失败，颜色还是可以正常的

                            //非完成情况下，关闭，属于 中断部署 状态，此处用于停止 下载 office 进程
                            Current_StageType = ProcessStage.Interrupt;
                            /*Pointing(ProcessStage.Interrupt); 暂停中断打点 */   //中断 点位。下载时触发该逻辑，打点会失败。
                        }
                        else
                        {
                            //完成状态时，清理文件夹
                            Lib_AppSdk.Clean();     //清理SDK目录
                        }

                        return false;
                    }
                }
            }
        }
    }
}
