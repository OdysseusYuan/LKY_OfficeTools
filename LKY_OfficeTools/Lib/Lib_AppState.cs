/*
 *      [LKY Office Tools] Copyright (C) 2022 - 2024 LiuKaiyuan Inc.
 *      
 *      FileName : Lib_AppState.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

namespace LKY_OfficeTools.Lib
{
    internal class Lib_AppState
    {
        internal enum RunMode
        {
            Manual,

            Passive,

            Service
        }

        internal static RunMode Current_RunMode = RunMode.Manual;

        internal enum ProcessStage
        {
            Starting = 1,

            Process = 2,

            Update_Success = 4,

            Update_Fail = 8,

            Interrupt = 16,

            RestartPC = 32,

            Finish_Success = 64,

            Finish_Fail = 128,
        }

        internal static ProcessStage Current_StageType = ProcessStage.Process;

        internal static bool Must_Use_PersonalDir
        {
            get; set;
        }
    }
}
