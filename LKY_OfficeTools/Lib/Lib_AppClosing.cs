/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppClosing.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Runtime.InteropServices;
using static LKY_OfficeTools.Common.Com_ExeOS.KillExe;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// APP 关闭时的类库
    /// </summary>
    internal class Lib_AppClosing
    {
        /// <summary>
        /// 关闭窗口的类库
        /// </summary>
        internal class CloseWindow
        {
            internal delegate bool ControlCtrlDelegate(int CtrlType);

            [DllImport("kernel32.dll")]
            internal static extern bool SetConsoleCtrlHandler(ControlCtrlDelegate HandlerRoutine, bool Add);

            internal static ControlCtrlDelegate newDelegate = new ControlCtrlDelegate(HandlerRoutine);

            /// <summary>
            /// 用户关闭窗口事件
            /// </summary>
            /// <param name="CtrlType"></param>
            /// <returns></returns>
            internal static bool HandlerRoutine(int CtrlType)
            {
                //只要程序不是已完成（无论成功与否），手动关闭，就会显示文字并打点
                if (Current_StageType != ProcessStage.Finish_Fail && Current_StageType != ProcessStage.Finish_Success)
                {
                    new Log($"\n     × 正在尝试 取消部署，请稍候 ...", ConsoleColor.Red);

                    Current_StageType = ProcessStage.Interrupt;         //设置中断状态。非完成情况下，关闭，属于 中断部署 状态，此处用于停止 下载 office 进程
                    Console.ForegroundColor = ConsoleColor.Gray;        //重置颜色，如果第一次失败，颜色还是可以正常的

                    //结束残存进程
                    Lib_AppSdk.KillAllSdkProcess(KillMode.Try_Friendly);

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
