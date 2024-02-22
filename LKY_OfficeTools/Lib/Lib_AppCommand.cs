/*
 *      [LKY Office Tools] Copyright (C) 2022 - 2024 LiuKaiyuan Inc.
 *      
 *      FileName : Lib_AppCommand.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_AppCommand
    {
        internal static ArgsFlag AppCommandFlag { get; set; }

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
                            //对于 /passive 模式，自动跳过所有需要确认的步骤
                            if (now_arg.Contains("/passive"))
                            {
                                SkipAllConfirm();
                                Current_RunMode = RunMode.Passive;  //设置为被动模式
                                break;
                            }
                            //服务模式。该模式 与 passive 是互斥的。
                            else if (now_arg.Contains("/service"))
                            {
                                Current_RunMode = RunMode.Service;  //设置为服务模式

                                //以服务模式运行
                                Lib_AppServiceConfig.Start();       //在Hub类库设置中，service模式将被添加passive标记。

                                //找到服务模式，不再执行后续的
                                Environment.Exit(-100);
                            }
                            //非服务、非被动模式
                            else
                            {
                                Current_RunMode = RunMode.Manual;   //设置为手动模式

                                if (now_arg.ToLower().Contains("/none_welcome_confirm"))
                                {
                                    AppCommandFlag |= ArgsFlag.None_Welcome_Confirm;
                                }

                                if (now_arg.ToLower().Contains("/ignore_manual_update_msg"))
                                {
                                    AppCommandFlag |= ArgsFlag.Ignore_Manual_Update_Msg;
                                }

                                if (now_arg.ToLower().Contains("/auto_remove_conflict_office"))
                                {
                                    AppCommandFlag |= ArgsFlag.Auto_Remove_Conflict_Office;
                                }

                                if (now_arg.ToLower().Contains("/none_finish_presskey"))
                                {
                                    AppCommandFlag |= ArgsFlag.None_Finish_PressKey;
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

        [Flags]
        internal enum ArgsFlag
        {
            None_Welcome_Confirm = 2,

            Ignore_Manual_Update_Msg = 4,

            Auto_Remove_Conflict_Office = 8,

            None_Finish_PressKey = 16,
        }

        private static bool SkipAllConfirm()
        {
            try
            {
                AppCommandFlag |=
                    ArgsFlag.None_Welcome_Confirm |
                    ArgsFlag.Auto_Remove_Conflict_Office |
                    ArgsFlag.Ignore_Manual_Update_Msg |
                    ArgsFlag.None_Finish_PressKey;
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
