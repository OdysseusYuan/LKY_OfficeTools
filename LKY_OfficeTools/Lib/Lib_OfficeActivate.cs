/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeActivate.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
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
            string info = Lib_SelfUpdate.latest_info;
            if (string.IsNullOrEmpty(info))
            {
                info = Com_WebOS.Visit_WebClient(Lib_SelfUpdate.update_json_url);
            }

            string KMS_info = Com_TextOS.GetCenterText(info, "\"KMS_List\": \"", "\"");

            //为空抛出异常
            if (!string.IsNullOrEmpty(KMS_info))
            {
                KMS_List = new List<string>(KMS_info.Split(','));
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

                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine($"\n------> 正在激活 Microsoft Office v{OfficeNetVersion.latest_version} ...");

                //执行：安装序列号
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"\n     >> 安装 Office 序列号 ...");
                string log_install_key = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_install_key})");
                if (!log_install_key.ToLower().Contains("successful"))
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"     × 安装序列号失败，激活终止，请稍后重试！如有问题请联系开发者。");
                    return -2;
                }
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ 安装序列号完成。");

                //执行：设置激活KMS地址
                string kms_flag = kms_server.Replace("kms.", "");
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"\n     >> 设置 Office [{kms_flag}] 激活载体 ...");
                string log_kms_url = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_kms_url})");
                if (!log_kms_url.ToLower().Contains("successful"))
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"     × 设置激活载体失败，激活终止，请稍后重试！如有问题请联系开发者。");
                    return 0;
                }
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ 设置激活载体完成。");

                //执行：开始激活
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"\n     >> 执行 Office 激活 ...");
                string log_activate = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_activate})");
                if (!log_activate.ToLower().Contains("successful"))
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"     × 无法执行激活，请稍后重试！如有问题请联系开发者。");
                    return -1;
                }
                /*Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"     √ 执行激活完成。");*/

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ Microsoft Office v{OfficeNetVersion.latest_version} 激活成功。");

                return 1;

                /*
                //汇总一句话命令
                string cmd = $"({cmd_switch_cd})&({cmd_install_key})&({cmd_kms_url})&({cmd_activate})";

                //开始激活
                string log_activate = Com_ExeOS.RunCmd(cmd);
                Console.WriteLine("结果：" + log_activate);
                */
            }
            return -10;
        }
    }
}
