/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeActivate.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
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
        /// 激活 Office
        /// </summary>
        internal static void Activating()
        {
            //检查安装情况
            OfficeLocalInstall.State install_state = OfficeLocalInstall.GetState();
            if (install_state != OfficeLocalInstall.State.Nothing)
            {
                //只要安装了 Office 相对新一点的版本，就用KMS开始激活
                string cmd_switch_cd = $"pushd \"{Environment.CurrentDirectory + @"\SDK\Activate"}\"";      //切换至OSPP文件目录
                string cmd_install_key = "cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH";          //安装序列号，默认是 ProPlus2021VL 的
                string cmd_kms_url = "cscript ospp.vbs /sethst:kms.chinancce.com";                          //设置激活KMS地址
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
                    return;
                }
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ 安装序列号完成。");

                //执行：设置激活KMS地址
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"\n     >> 设置 Office KMS 激活服务器 ...");
                string log_kms_url = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_kms_url})");
                if (!log_kms_url.ToLower().Contains("successful"))
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"     × 设置激活KMS地址失败，激活终止，请稍后重试！如有问题请联系开发者。");
                    return;
                }
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ 设置激活KMS地址完成。");

                //执行：开始激活
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"\n     >> 执行 Office 激活 ...");
                string log_activate = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_activate})");
                if (!log_activate.ToLower().Contains("successful"))
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"     × 无法执行激活，请稍后重试！如有问题请联系开发者。");
                    return;
                }
                /*Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"     √ 执行激活完成。");*/

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ Microsoft Office v{OfficeNetVersion.latest_version} 激活成功。");

                /*
                //汇总一句话命令
                string cmd = $"({cmd_switch_cd})&({cmd_install_key})&({cmd_kms_url})&({cmd_activate})";

                //开始激活
                string log_activate = Com_ExeOS.RunCmd(cmd);
                Console.WriteLine("结果：" + log_activate);
                */
            }
        }
    }
}
