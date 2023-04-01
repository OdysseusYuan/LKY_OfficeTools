/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2023 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppState.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

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
            /// 升级成功
            /// </summary>
            Update_Success = 8, 
            
            /// <summary>
            /// 升级失败
            /// </summary>
            Update_Fail = 16,

            /// <summary>
            /// 强制结束，中断运行
            /// </summary>
            Interrupt = 32,

            /// <summary>
            /// 已结束，且运行成功。即使存在等待用户输入，但程序自身已经完成
            /// </summary>
            Finish_Success = 64,

            /// <summary>
            /// 已结束，运行有严重错误。导致失败。即使存在等待用户输入，但程序自身已经完成
            /// </summary>
            Finish_Fail = 128,
        }

        /// <summary>
        /// APP 当前状态（初始值为 Process）
        /// </summary>
        internal static ProcessStage Current_StageType = ProcessStage.Process;

        /// <summary>
        /// 是否必须使用个人文档目录？
        /// 通常情况，为False，使用CommonApplicationData，用于方便部署服务。
        /// 但有些电脑因权限问题，无法使用该目录，设置为true后，AppInfo的Documents_Root将变成个人目录。
        /// </summary>
        internal static bool Must_Use_PersonalDir
        {
            get; set;
        }
    }
}
