/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppState.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Runtime.InteropServices;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// APP 运行状态 类库
    /// </summary>
    internal class Lib_AppState
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
            /// 被动模式运行
            /// </summary>
            Passive,

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
        /// APP 运行状态类型
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
