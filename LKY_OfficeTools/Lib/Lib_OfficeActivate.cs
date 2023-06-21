/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeActivate.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.AppPath;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo.OfficeLocalInfo;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 实现激活的类库
    /// </summary>
    internal class Lib_OfficeActivate
    {
        /// <summary>
        /// KMS 服务器列表
        /// </summary>
        internal static List<string> KMS_List = new List<string>();

        /// <summary>
        /// 获取 KMS 服务器列表
        /// </summary>
        internal static void Activating()
        {
            string KMS_info = Com_TextOS.GetCenterText(AppJson.Info, "\"KMS_List\": \"", "\"");

            //为空抛出异常
            if (!string.IsNullOrEmpty(KMS_info))
            {
                int try_times = 1;                                  //激活尝试的次数，初始值为1
                KMS_List = new List<string>(KMS_info.Split(';'));
                foreach (var now_kms in KMS_List)
                {
                    //激活成功时，结束；未安装Office导致不成功，也跳出。其余问题多次尝试不同激活服务器
                    int act_state = StartActivate(now_kms.Replace(" ", ""));   //替换空格并激活
                    if (act_state == 1 || act_state < -2)
                    {
                        //激活成功（1），或者安装本身存在问题（< -11），亦或者安装序列号本身有问题（-3），直接结束激活。
                        break;
                    }
                    else
                    {
                        if (try_times < KMS_List.Count)
                        {
                            new Log($"\n     >> 即将尝试第 {++try_times} 次激活 ...", ConsoleColor.DarkYellow);
                        }
                        continue;
                    }
                }
            }
            else
            {
                //获取失败时，使用默认值
                StartActivate();
            }
        }

        /// <summary>
        /// 激活 Office
        /// 没安装 Office 最新版 返回 -10，激活成功 = 1，其余小于1的值均为失败
        /// </summary>
        internal static int StartActivate(string kms_server = "kms.chinancce.com")
        {
            //检查安装情况
            InstallState install_state = GetOfficeState();
            if (install_state == InstallState.Correct)              //必须安装最新版，才能激活
            {
                //检查 ospp.vbs 文件是否存在
                if (!File.Exists(Documents.SDKs.Activate_OSPP))
                {
                    new Log($"     × 目录 {Documents.SDKs.Activate} 下文件丢失！", ConsoleColor.DarkRed);
                    return -4;
                }

                //只要安装了 Office 新版本，就用KMS开始激活
                string cmd_switch_cd = $"pushd \"{Documents.SDKs.Activate}\"";          //切换至OSPP文件目录
                string cmd_kms_url = $"cscript ospp.vbs /sethst:{kms_server}";                          //设置激活KMS地址
                string cmd_activate = "cscript ospp.vbs /act";                                              //开始激活

                new Log($"\n------> 正在激活 Office v{OfficeNetInfo.OfficeLatestVersion} ...", ConsoleColor.DarkCyan);

                //执行：设置激活KMS地址
                string kms_flag = kms_server.Replace("kms.", "");
                new Log($"\n     >> 设置 Office [{kms_flag}] 激活载体 ...", ConsoleColor.DarkYellow);
                string log_kms_url = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_kms_url})");
                if (!log_kms_url.ToLower().Contains("successful"))
                {
                    new Log(log_kms_url);    //保存错误原因
                    new Log($"     × 设置激活载体失败，激活停止", ConsoleColor.DarkRed);
                    return -2;
                }
                new Log($"     √ 已完成 Office 激活载体设置。", ConsoleColor.DarkGreen);

                //执行：开始激活
                new Log($"\n     >> 执行 Office 激活 ...", ConsoleColor.DarkYellow);
                string log_activate = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_activate})");

                //先判断是几个SKU项目，以及成功数量
                int sku_count = Com_TextOS.GetStringTimes(log_activate.ToLower(), "sku id");
                //获取成功的数量
                int success_count = Com_TextOS.GetStringTimes(log_activate.ToLower(), "successful");

                bool activate_success;      //激活成功标志
                if (success_count > 0 & sku_count == success_count)
                {
                    //全部激活成功
                    activate_success = true;
                }
                else
                {
                    //至少有1个激活失败
                    activate_success = false;
                    new Log($"     × 有 {sku_count - success_count} 个（共 {sku_count} 个）产品架构未能成功激活。", ConsoleColor.DarkRed);
                }

                //判断原因
                if (!activate_success)
                {
                    //继续判断失败原因，并给出方案

                    //0x80080005
                    if (log_activate.Contains("0x80080005"))
                    {
                        //0x80080005错误：劫持问题，自动修复
                        new Log($"     >> 尝试修复 0x80080005 问题中 ...", ConsoleColor.DarkYellow);
                        string base_reg = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options";
                        string spp_reg = "SppExtComObj.exe";

                        //清除x32劫持
                        var x32_spp = Register.ExistItem(RegistryHive.LocalMachine, RegistryView.Registry32, $@"{base_reg}\{spp_reg}");
                        if (x32_spp)
                        {
                            Register.DeleteItem(RegistryHive.LocalMachine, RegistryView.Registry32, base_reg, spp_reg);
                        }

                        //清除x64劫持
                        var x64_spp = Register.ExistItem(RegistryHive.LocalMachine, RegistryView.Registry64, $@"{base_reg}\{spp_reg}");
                        if (x64_spp)
                        {
                            Register.DeleteItem(RegistryHive.LocalMachine, RegistryView.Registry64, base_reg, spp_reg);
                        }

                        new Log($"     √ 已完成 0x80080005 修复，稍后将自动重试激活。", ConsoleColor.DarkGreen);
                    }
                    //0x8007000D
                    else if (log_activate.Contains("0x8007000D"))
                    {
                        //0x8007000D错误：软件保护、日期时间问题，自动修复
                        new Log($"     >> 尝试修复 0x8007000D 问题中，请同时确保您的计算机 日期/时间 正确 ...", ConsoleColor.DarkYellow);

                        //--------------------------------------- 软件保护修复 ---------------------------------------
                        //先停止软件保护服务（sppsvc）
                        Com_ServiceOS.Action.Stop("sppsvc");    //无论是否成功都继续

                        //准备清除注册表路径
                        string base_reg = @"SOFTWARE\Microsoft";
                        string sub_reg = "OfficeSoftwareProtectionPlatform";

                        //清除x32
                        var x32_spp = Register.ExistItem(RegistryHive.LocalMachine, RegistryView.Registry32, $@"{base_reg}\{sub_reg}");
                        if (x32_spp)
                        {
                            Register.DeleteItem(RegistryHive.LocalMachine, RegistryView.Registry32, base_reg, sub_reg);
                        }

                        //清除x64
                        var x64_spp = Register.ExistItem(RegistryHive.LocalMachine, RegistryView.Registry64, $@"{base_reg}\{sub_reg}");
                        if (x64_spp)
                        {
                            Register.DeleteItem(RegistryHive.LocalMachine, RegistryView.Registry64, base_reg, sub_reg);
                        }

                        //再次启动软件保护服务（sppsvc）
                        Com_ServiceOS.Action.Start("sppsvc");   //无论是否成功都继续

                        //--------------------------------------- 软件保护修复（完成） ---------------------------------------


                        //--------------------------------------- 系统日期/时间/时区修复 ---------------------------------------

                        //--------------------------------------- 系统日期/时间/时区修复（完成） ---------------------------------------

                        new Log($"     √ 已完成 0x8007000D 修复，稍后将自动重试激活。", ConsoleColor.DarkGreen);
                    }
                    //0x80040154
                    else if (log_activate.Contains("0x80040154"))
                    {
                        //0x80040154错误：没有注册类
                        new Log($"     × 系统可能存在损坏，建议您重新安装操作系统后重试！", ConsoleColor.DarkRed);
                        return -101;    //返回无限小，不再重试
                    }
                    //0xC004F074
                    else if (log_activate.Contains("0xC004F074"))
                    {
                        //0xC004F074错误：与KMS服务器通讯失败
                        new Log($"     × 激活失败！若此消息频频复现，强烈建议您重置网卡设置 或 重新安装操作系统！", ConsoleColor.DarkRed);
                    }
                    else
                    {
                        //非已知问题
                        new Log(log_activate);    //保存错误原因
                        new Log($"     × 意外的错误导致激活失败！", ConsoleColor.DarkRed);
                    }

                    return -1;
                }

                new Log($"     √ 已完成 Office v{OfficeNetInfo.OfficeLatestVersion} 正版激活。", ConsoleColor.DarkGreen);
                Lib_AppState.Current_StageType = Lib_AppState.ProcessStage.Finish_Success;   //设置整体运行状态为成功

                return 1;
            }
            else if (install_state == InstallState.Diff)
            {
                new Log($"     × 当前系统未安装最新版本的 Office，激活停止！", ConsoleColor.DarkRed);
                return -12;
            }
            else if (install_state == InstallState.Multi)
            {
                new Log($"     × 当前系统存在多个 Office 版本，无法完成激活！", ConsoleColor.DarkRed);    //这种多版本出错是指，未正确安装最新版，而且系统还有多个版本
                return -14;
            }
            else if (install_state == InstallState.None)
            {
                new Log($"     × 当前系统未安装任何 Office 版本，不需要激活！", ConsoleColor.DarkRed);
                return -18;
            }
            else
            {
                new Log($"     × 因其它问题，Office 激活被迫停止！", ConsoleColor.DarkRed);
                return -99;
            }
        }
    }
}
