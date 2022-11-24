/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeActivate.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.IO;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.AppPath;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo.OfficeLocalInstall;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 实现激活的类库
    /// </summary>
    internal class Lib_OfficeActivate
    {
        /// <summary>
        /// 重载实现激活
        /// </summary>
        internal Lib_OfficeActivate()
        {
            Activating();
        }

        /// <summary>
        /// KMS 服务器列表
        /// </summary>
        internal static List<string> KMS_List = new List<string>();

        /// <summary>
        /// 获取 KMS 服务器列表
        /// </summary>
        internal static void Activating()
        {
            //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
            string info = Lib_AppUpdate.latest_info;
            if (string.IsNullOrEmpty(info))
            {
                info = Com_WebOS.Visit_WebClient(Lib_AppUpdate.update_json_url);
            }

            string KMS_info = Com_TextOS.GetCenterText(info, "\"KMS_List\": \"", "\"");

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
            if (install_state.HasFlag(InstallState.Correct))      //只要安装了最新版，无论是否有多重版本叠加安装，均尝试激活
            {
                //检查 ospp.vbs 文件是否存在
                if (!File.Exists(Documents.SDKs.Activate_OSPP))
                {
                    new Log($"     × 目录 {Documents.SDKs.Activate} 下文件丢失！", ConsoleColor.DarkRed);
                    return -4;
                }

                //只要安装了 Office 新版本，就用KMS开始激活
                string cmd_switch_cd = $"pushd \"{Documents.SDKs.Activate}\"";          //切换至OSPP文件目录
                string cmd_install_key = "cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH";          //安装序列号，默认是 ProPlus2021VL 的
                string cmd_kms_url = $"cscript ospp.vbs /sethst:{kms_server}";                          //设置激活KMS地址
                string cmd_activate = "cscript ospp.vbs /act";                                              //开始激活

                new Log($"\n------> 正在激活 Office v{OfficeNetVersion.latest_version} ...", ConsoleColor.DarkCyan);

                //执行：安装序列号
                new Log($"\n     >> 安装 Office 序列号 ...", ConsoleColor.DarkYellow);
                string log_install_key = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_install_key})");
                if (!log_install_key.ToLower().Contains("successful"))
                {
                    new Log(log_install_key);    //保存错误原因
                    new Log($"     × 安装序列号失败，激活停止。", ConsoleColor.DarkRed);
                    return -3;
                }
                new Log($"     √ 安装序列号完成。", ConsoleColor.DarkGreen);

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
                new Log($"     √ 设置激活载体完成。", ConsoleColor.DarkGreen);

                //执行：开始激活
                new Log($"\n     >> 执行 Office 激活 ...", ConsoleColor.DarkYellow);
                string log_activate = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_activate})");
                if (!log_activate.ToLower().Contains("successful"))
                {
                    new Log(log_activate);    //保存错误原因
                    new Log($"     × 无法执行激活，激活停止。", ConsoleColor.DarkRed);
                    return -1;
                }

                new Log($"     √ Office v{OfficeNetVersion.latest_version} 激活成功。", ConsoleColor.DarkGreen);
                Lib_AppState.Current_StageType = Lib_AppState.ProcessStage.Finish_Success;   //设置整体运行状态为成功

                return 1;
            }
            else if (install_state.HasFlag(InstallState.Diff))
            {
                new Log($"     × 当前系统未安装最新版本的 Office，激活停止！", ConsoleColor.DarkRed);
                return -12;
            }
            else if (install_state.HasFlag(InstallState.Multi))
            {
                new Log($"     × 当前系统存在多个 Office 版本，无法完成激活！", ConsoleColor.DarkRed);    //这种多版本出错是指，未正确安装最新版，而且系统还有多个版本
                return -14;
            }
            else if (install_state.HasFlag(InstallState.None))
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
