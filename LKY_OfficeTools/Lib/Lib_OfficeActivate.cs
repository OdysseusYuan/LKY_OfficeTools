/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeActivate.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppLog.Log;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;

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
                    if (act_state == 1 || act_state == -10)
                    {
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
        /// 没安装 Office 返回 -10，激活成功 = 1，其余小于1的值均为失败
        /// </summary>
        internal static int StartActivate(string kms_server = "kms.chinancce.com")
        {
            //检查安装情况
            OfficeLocalInstall.State install_state = OfficeLocalInstall.GetState();
            if (install_state != OfficeLocalInstall.State.Nothing)
            {
                //只要安装了 Office 相对新一点的版本，就用KMS开始激活
                string cmd_switch_cd = $"pushd \"{Environment.CurrentDirectory + @"\SDK\Activate"}\"";      //切换至OSPP文件目录
                string cmd_install_key = "cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH";          //安装序列号，默认是 ProPlus2021VL 的
                string cmd_kms_url = $"cscript ospp.vbs /sethst:{kms_server}";                          //设置激活KMS地址
                string cmd_activate = "cscript ospp.vbs /act";                                              //开始激活

                new Log($"\n------> 正在激活 Microsoft Office v{OfficeNetVersion.latest_version} ...", ConsoleColor.DarkCyan);

                //执行：安装序列号
                new Log($"\n     >> 安装 Office 序列号 ...", ConsoleColor.DarkYellow);
                string log_install_key = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_install_key})");
                if (!log_install_key.ToLower().Contains("successful"))
                {
                    new Log(log_install_key);    //保存错误原因
                    new Log($"     × 安装序列号失败，激活停止。", ConsoleColor.DarkRed);
                    return -2;
                }
                new Log($"     √ 安装序列号完成。", ConsoleColor.DarkGreen);

                //执行：设置激活KMS地址
                string kms_flag = kms_server.Replace("kms.", "");
                new Log($"\n     >> 设置 Office [{kms_flag}] 激活载体 ...", ConsoleColor.DarkYellow);
                string log_kms_url = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_kms_url})");
                if (!log_kms_url.ToLower().Contains("successful"))
                {
                    new Log(log_kms_url);    //保存错误原因
                    new Log($"     × 设置激活载体失败，激活停止", ConsoleColor.DarkRed);
                    return 0;
                }
                new Log($"     √ 设置激活载体完成。", ConsoleColor.DarkGreen);

                //执行：开始激活
                new Log($"\n     >> 执行 Office 激活 ...", ConsoleColor.DarkYellow);
                string log_activate = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_activate})");
                if (!log_activate.ToLower().Contains("successful"))
                {
                    new Log(log_activate);    //保存错误原因
                    new Log($"     × 无法执行激活，激活停止。", ConsoleColor.DarkRed);
                    return -1;
                }
                /*Console.ForegroundColor = ConsoleColor.Green;
                new Log($"     √ 执行激活完成。");*/

                new Log($"     √ Microsoft Office v{OfficeNetVersion.latest_version} 激活成功。", ConsoleColor.DarkGreen);

                return 1;

                /*
                //汇总一句话命令
                string cmd = $"({cmd_switch_cd})&({cmd_install_key})&({cmd_kms_url})&({cmd_activate})";

                //开始激活
                string log_activate = Com_ExeOS.RunCmd(cmd);
                new Log("结果：" + log_activate);
                */
            }
            return -10;
        }
    }
}
