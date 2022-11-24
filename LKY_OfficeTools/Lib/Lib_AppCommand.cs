/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppCommand.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Windows.Forms;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// App 命令行类库
    /// </summary>
    internal class Lib_AppCommand
    {
        /// <summary>
        /// 当前 APP 启动时的命令行
        /// </summary>
        internal static ArgsFlag AppCommandFlag { get; set; }

        /// <summary>
        /// 重载获得命令行
        /// </summary>
        internal Lib_AppCommand(string[] args)
        {
            try
            {
                //非空判断
                if (args != null && args.Length > 0)
                {
                    foreach (var now_arg in args)
                    {
                        //单独arg非空判断
                        if (!string.IsNullOrWhiteSpace(now_arg))
                        {
                            //对于服务模式，或者手动指定的 /passive 模式，自动跳过所有需要确认的步骤
                            if (now_arg.Contains("/service") || now_arg.Contains("/passive"))
                            {
                                SkipAllConfirm();
                                App.State.Current_RunMode = App.State.RunMode.Service;
                                break;
                            }
                            else
                            {
                                if (now_arg.Contains("/none_welcome_confirm"))
                                {
                                    AppCommandFlag |= ArgsFlag.None_Welcome_Confirm;
                                }

                                if (now_arg.Contains("/none_finish_presskey"))
                                {
                                    AppCommandFlag |= ArgsFlag.None_Finish_PressKey;
                                }

                                if (now_arg.Contains("/auto_remove_conflict_office"))
                                {
                                    AppCommandFlag |= ArgsFlag.Auto_Remove_Conflict_Office;
                                }

                                continue;
                            }
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }

        /// <summary>
        /// 命令行标记
        /// </summary>
        [Flags]
        internal enum ArgsFlag
        {
            /// <summary>
            /// 无需欢迎界面确认，自动继续
            /// </summary>
            None_Welcome_Confirm = 2,

            /// <summary>
            /// 无需最终完成界面的“请按任意键退出”
            /// </summary>
            None_Finish_PressKey = 4,

            /// <summary>
            /// 在安装中，自动移除冲突的 Office 版本
            /// </summary>
            Auto_Remove_Conflict_Office = 8,
        }

        /// <summary>
        /// 跳过所有确认步骤
        /// </summary>
        /// <returns></returns>
        private static bool SkipAllConfirm()
        {
            try
            {
                AppCommandFlag |= 
                    ArgsFlag.None_Welcome_Confirm |
                    ArgsFlag.None_Finish_PressKey |
                    ArgsFlag.Auto_Remove_Conflict_Office;
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
