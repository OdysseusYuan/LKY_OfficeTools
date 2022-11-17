/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
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
            /// 路径类
            /// </summary>
            internal class Path
            {
                /// <summary>
                /// APP 文档根目录
                /// </summary>
                internal static string Dir_Document = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\LKY Office Tools";

                /// <summary>
                /// APP 日志存储目录
                /// </summary>
                internal static string Dir_Log = $"{Dir_Document}\\Logs";

                /// <summary>
                /// APP 临时文件夹目录
                /// </summary>
                internal static string Dir_Temp = $"{Dir_Document}\\Temp";

                /// <summary>
                /// APP SDK文件夹目录
                /// </summary>
                internal static string Dir_SDK = $"{Dir_Document}\\SDKs";
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
                /// APP运行状态类型
                /// </summary>
                internal enum RunType
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
                internal static RunType Current_Runtype = RunType.Process;

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
                        if (Current_Runtype != RunType.Finish_Fail && Current_Runtype != RunType.Finish_Success)
                        {
                            switch (CtrlType)
                            {
                                //Ctrl+C关闭
                                case 0:
                                    new Log($"\n     × 部署正在取消，请稍候 ...", ConsoleColor.DarkRed);
                                    Pointing(RunType.Interrupt);
                                    Lib_AppSdk.Clean();     //清理SDK目录
                                    break;

                                //按 控制台关闭按钮 关闭
                                case 2:
                                    new Log($"\n     × 部署正在中断，请稍候 ...", ConsoleColor.DarkRed);
                                    Pointing(RunType.Interrupt);
                                    Lib_AppSdk.Clean();     //清理SDK目录
                                    break;
                            }
                        }
                        return false;
                    }
                }
            }
        }
    }
}
