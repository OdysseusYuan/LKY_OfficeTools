/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInstall.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using System.Threading;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;
using static LKY_OfficeTools.Lib.Lib_SelfLog;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// Office 安装类库
    /// </summary>
    internal class Lib_OfficeInstall
    {
        /// <summary>
        /// 重载实现安装
        /// </summary>
        internal Lib_OfficeInstall()
        {
            //下载后，开始安装
            int DownCode = Lib_OfficeDownload.FilesDownload();

            //判断下载情况
            switch (DownCode)
            {
                case 1:
                    if (ConflictCheck())        //冲突检查，通过进入安装，不通过直接返回false
                    {
                        //安装成功，进入激活程序
                        new Lib_OfficeActivate();
                    }
                    else
                    {
                        new Log($"     × 因 Office 安装失败，自动跳过激活流程！", ConsoleColor.DarkRed);
                    }
                    return;
                case 0:
                    new Log($"     × 未能找到可用的 Office 安装文件！", ConsoleColor.DarkRed);
                    return;
                case -1:
                    //无需下载安装，直接进入激活模块
                    new Lib_OfficeActivate();
                    return;
            }
        }

        /// <summary>
        /// 冲突版本检查，卸载和主发行版本不一样的的 Office
        /// </summary>
        internal static bool ConflictCheck()
        {
            try
            {
                new Log($"\n------> 正在进行 Office 冲突检查 ...", ConsoleColor.DarkCyan);

                //先获取目前已经安装的 Office 版本
                string Current_Office_ID = Com_SystemOS.Registry.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                string info = Lib_SelfUpdate.latest_info;
                if (string.IsNullOrEmpty(info))
                {
                    info = Com_WebOS.Visit_WebClient(Lib_SelfUpdate.update_json_url);
                }
                string Pop_Office_ID = Com_TextOS.GetCenterText(info, "\"Pop_Office_ID\": \"", "\"");
                ///获取失败时，默认使用 2021VOL 版
                if (string.IsNullOrEmpty(Pop_Office_ID))
                {
                    Pop_Office_ID = "ProPlus2021Volume";
                }

                //Office ID完全不同时，需要卸载旧版本
                if (Current_Office_ID != Pop_Office_ID)
                {
                    new Log($"      * 发现冲突的 Office 版本：{Current_Office_ID}，如需安装最新版，请先卸载旧版本。", ConsoleColor.DarkRed);

                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.Write($"        卸载旧版全部组件，并仅安装新版 Word、Excel、PPT 三件套，请按 回车键 继续 ...");
                    if (Console.ReadKey().Key == ConsoleKey.Enter)
                    {
                        new Log($"\n     √ 您已确认卸载 Office [{Current_Office_ID}] 旧版本。", ConsoleColor.DarkGreen);

                        //定义ODT文件位置
                        string ODT_path_root = Environment.CurrentDirectory + @"\SDK\ODT\";
                        string ODT_path_exe = ODT_path_root + @"ODT.exe";
                        string ODT_path_xml = ODT_path_root + @"uninstall.xml";    //此文件需要新生成

                        //生成卸载xml
                        string xml_content = "<Configuration>\n  <Remove All=\"TRUE\" />\n  <Display Level=\"NONE\" AcceptEULA=\"TRUE\"/>\n</Configuration>";
                        File.WriteAllText(ODT_path_xml, xml_content);

                        //检查ODT文件是否存在
                        if (!File.Exists(ODT_path_exe) || !File.Exists(ODT_path_xml))
                        {
                            new Log($"     × 目录：{ODT_path_root} 下文件丢失，请重新下载本软件！", ConsoleColor.DarkRed);
                            return false;
                        }

                        //执行卸载命令
                        new Log($"\n------> 正在卸载 Office [{Current_Office_ID}] 旧版本 ...", ConsoleColor.DarkCyan);
                        Thread.Sleep(1000);     //基于体验，延迟1s
                        new Log($"     >> 此过程大约会持续几分钟的时间，请耐心等待 ...", ConsoleColor.DarkYellow);

                        string uninstall_args = $"/configure \"{ODT_path_xml}\"";
                        bool isUninstall = Com_ExeOS.RunExe(ODT_path_exe, uninstall_args);
                        var reg_info = Com_SystemOS.Registry.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                        //未正常结束卸载 OR 注册表键值不为空 时，视为卸载失败
                        if (!isUninstall || !string.IsNullOrEmpty(reg_info))
                        {
                            new Log($"     × 卸载冲突版本失败。您可联系开发者进行咨询！", ConsoleColor.DarkRed);
                            return false;
                        }

                        //卸载正常 + 注册表已清空时，开始安装新版本
                        new Log($"     √ 卸载 Office [{Current_Office_ID}] 旧版本完成。", ConsoleColor.DarkGreen);

                        //卸载成功后，开始安装新版本
                        return StartInstall();
                    }
                    else
                    {
                        new Log($"\n     × 您已拒绝卸载 Office 其他版本，新版本无法安装！", ConsoleColor.DarkRed);
                        return false;
                    }
                }
                else
                {
                    //不存在版本冲突时，直接开始安装
                    return StartInstall();
                }
            }
            catch
            {
                return false;
            }
        }

        /* ----------- 因 v1.0.3 及其以前的版本存在 升级时的 有新目录拷贝会导致异常的错误，
         * ----------- 当前迭代暂时不考虑使用 SaRAC
        /// <summary>
        /// 冲突版本检查，卸载和主发行版本不一样的的 Office
        /// </summary>
        internal static bool ConflictCheck()
        {
            try
            {
                new Log($"\n------> 正在进行 Office 冲突检查 ...", ConsoleColor.DarkCyan);

                //先获取目前已经安装的 Office 版本
                string Current_Office_ID = Com_SystemOS.Registry.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                string info = Lib_SelfUpdate.latest_info;
                if (string.IsNullOrEmpty(info))
                {
                    info = Com_WebOS.Visit_WebClient(Lib_SelfUpdate.update_json_url);
                }
                string Pop_Office_ID = Com_TextOS.GetCenterText(info, "\"Pop_Office_ID\": \"", "\"");
                ///获取失败时，默认使用 2021VOL 版
                if (string.IsNullOrEmpty(Pop_Office_ID))
                {
                    Pop_Office_ID = "ProPlus2021Volume";
                }

                //Office ID完全不同时，需要卸载旧版本
                if (Current_Office_ID != Pop_Office_ID)
                {
                    new Log($"      * 发现冲突的 Office 版本：{Current_Office_ID}，如需安装最新版，请先卸载旧版本。", ConsoleColor.DarkRed);

                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.Write($"        卸载旧版全部组件，并仅安装新版 Word、Excel、PPT 三件套，请按 回车键 继续 ...");
                    if (Console.ReadKey().Key == ConsoleKey.Enter)
                    {
                        new Log($"\n     √ 您已确认卸载 Office [{Current_Office_ID}] 旧版本。", ConsoleColor.DarkGreen);

                        //定义SaRAC文件位置
                        string SaRAC_path_root = Environment.CurrentDirectory + @"\SDK\SaRAC\";
                        string SaRAC_path_exe = SaRAC_path_root + @"SaRACmd.exe";

                        //检查SaRAC文件是否存在
                        if (!File.Exists(SaRAC_path_exe))
                        {
                            new Log($"     × 目录：{SaRAC_path_root} 下文件丢失，请重新下载本软件！", ConsoleColor.DarkRed);
                            return false;
                        }

                        //执行卸载命令
                        new Log($"\n------> 正在卸载 Office [{Current_Office_ID}] 旧版本 ...", ConsoleColor.DarkCyan);
                        Thread.Sleep(1000);     //基于体验，延迟1s
                        new Log($"     >> 此过程大约会持续几分钟的时间，请耐心等待 ...", ConsoleColor.DarkYellow);

                        string cmd_switch_cd = $"pushd \"{SaRAC_path_root}\"";             //切换至SaRAC文件目录
                        string cmd_unstall = $"SaRAcmd.exe -S OfficeScrubScenario -AcceptEula -Officeversion All";
                        string uninstall_result = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_unstall})");
                        if (!uninstall_result.ToLower().Contains("successful"))
                        {
                            //首次卸载不成功时，尝试 ODT 模式卸载
                            //定义ODT文件位置
                            string ODT_path_root = Environment.CurrentDirectory + @"\SDK\ODT\";
                            string ODT_path_exe = ODT_path_root + @"ODT.exe";
                            string ODT_path_xml = ODT_path_root + @"uninstall.xml";    //此文件需要新生成

                            //生成卸载xml
                            string xml_content = "<Configuration>\n  <Remove All=\"TRUE\" />\n  <Display Level=\"NONE\" AcceptEULA=\"TRUE\"/>\n</Configuration>";
                            File.WriteAllText(ODT_path_xml, xml_content);

                            //检查ODT文件是否存在
                            if (!File.Exists(ODT_path_exe) || !File.Exists(ODT_path_xml))
                            {
                                new Log($"     × 目录：{ODT_path_root} 下文件丢失，请重新下载本软件！", ConsoleColor.DarkRed);
                                return false;
                            }

                            //执行卸载命令
                            new Log($"     >> 再次尝试卸载 Office [{Current_Office_ID}] 旧版本 ...", ConsoleColor.DarkCyan);
                            string uninstall_args = $"/configure \"{ODT_path_xml}\"";
                            bool isUninstall = Com_ExeOS.RunExe(ODT_path_exe, uninstall_args);
                            var reg_info = Com_SystemOS.Registry.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");
                            
                            //未正常结束卸载 OR 注册表键值不为空 时，视为卸载失败
                            if (!isUninstall || !string.IsNullOrEmpty(reg_info))
                            {
                                new Log($"     × 卸载冲突版本失败。您可联系开发者进行咨询！", ConsoleColor.DarkRed);
                                return false;
                            }
                        }
                        new Log($"     √ 卸载 Office [{Current_Office_ID}] 旧版本完成。", ConsoleColor.DarkGreen);

                        //卸载成功后，开始安装新版本
                        return StartInstall();
                    }
                    else
                    {
                        new Log($"\n     × 您已拒绝卸载 Office 其他版本，新版本无法安装！", ConsoleColor.DarkRed);
                        return false;
                    }
                }
                else
                {
                    //不存在版本冲突时，直接开始安装
                    return StartInstall();
                }
            }
            catch
            {
                return false;
            }
        }*/

        /// <summary>
        /// 开始安装 Office
        /// </summary>
        internal static bool StartInstall()
        {
            //定义ODT文件位置
            string ODT_path_root = Environment.CurrentDirectory + @"\SDK\ODT\";
            string ODT_path_exe = ODT_path_root + @"ODT.exe";
            string ODT_path_xml = ODT_path_root + @"config.xml";

            //检查ODT文件是否存在
            if (!File.Exists(ODT_path_exe) || !File.Exists(ODT_path_xml))
            {
                new Log($"     × 目录：{ODT_path_root} 下文件丢失，请重新下载本软件！", ConsoleColor.DarkRed);
                return false;
            }

            //修改新的xml信息
            ///修改安装目录，安装目录为运行根目录
            bool isNewInstallPath = Com_FileOS.XML.SetValue(ODT_path_xml, "SourcePath", Environment.CurrentDirectory);

            //检查是否修改成功（安装目录）
            if (!isNewInstallPath)
            {
                new Log($"     × 配置 Install 信息错误！", ConsoleColor.DarkRed);
                return false;
            }

            ///修改为新版本号
            bool isNewVersion = Com_FileOS.XML.SetValue(ODT_path_xml, "Version", OfficeNetVersion.latest_version.ToString());

            //检查是否修改成功（版本号）
            if (!isNewVersion)
            {
                new Log($"     × 配置 Version 信息错误！", ConsoleColor.DarkRed);
                return false;
            }

            ///修改安装的位数
            //获取系统位数
            int sys_bit;
            if (Environment.Is64BitOperatingSystem)
            {
                sys_bit = 64;
            }
            else
            {
                sys_bit = 32;
            }
            bool isNewBit = Com_FileOS.XML.SetValue(ODT_path_xml, "OfficeClientEdition", sys_bit.ToString());

            //检查是否修改成功（位数）
            if (!isNewBit)
            {
                new Log($"     × 配置 Edition 信息错误！", ConsoleColor.DarkRed);
                return false;
            }

            ///修改 Product ID
            ///先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
            string info = Lib_SelfUpdate.latest_info;
            if (string.IsNullOrEmpty(info))
            {
                info = Com_WebOS.Visit_WebClient(Lib_SelfUpdate.update_json_url);
            }
            string Pop_Office_ID = Com_TextOS.GetCenterText(info, "\"Pop_Office_ID\": \"", "\"");
            ///获取失败时，默认使用 2021VOL 版
            if (string.IsNullOrEmpty(Pop_Office_ID))
            {
                Pop_Office_ID = "ProPlus2021Volume";
            }
            bool isNewID = Com_FileOS.XML.SetValue(ODT_path_xml, "Product ID", Pop_Office_ID);

            //检查是否修改成功（Product ID）
            if (!isNewID)
            {
                new Log($"     × 配置 Product ID 信息错误！", ConsoleColor.DarkRed);
                return false;
            }

            //开始安装
            string install_args = $"/configure \"{ODT_path_xml}\"";     //配置命令行

            new Log($"\n------> 开始安装 Microsoft Office v{OfficeNetVersion.latest_version} ...", ConsoleColor.DarkCyan);

            bool isInstallFinish = Com_ExeOS.RunExe(ODT_path_exe, install_args);

            //检查是否因配置不正确等导致，意外退出安装
            if (!isInstallFinish)
            {
                new Log($"     × Microsoft Office v{OfficeNetVersion.latest_version} 安装意外结束！", ConsoleColor.DarkRed);
                return false;
            }

            //检查安装是否成功
            OfficeLocalInstall.State install_state = OfficeLocalInstall.GetState();
            if (install_state == OfficeLocalInstall.State.Nothing)
            {
                //找不到 ClickToRun 注册表
                new Log($"     × Microsoft Office v{OfficeNetVersion.latest_version} 安装失败！", ConsoleColor.DarkRed);
                return false;
            }
            else
            {
                if (install_state == OfficeLocalInstall.State.Installed)
                {
                    //一切正常
                    Com_ProcessOS.KillProcess("OfficeClickToRun");      //结束无关进程
                    Com_ProcessOS.KillProcess("OfficeC2RClient");       //结束无关进程
                    new Log($"     √ Microsoft Office v{OfficeNetVersion.latest_version} 已安装完成。", ConsoleColor.DarkGreen);
                    return true;
                }
                else if (install_state == OfficeLocalInstall.State.VersionDiff)
                {
                    //版本号和一开始下载的版本号不一致
                    new Log($"     × 未能正确安装 Microsoft Office v{OfficeNetVersion.latest_version} 版本！", ConsoleColor.DarkGreen);
                    return false;
                }
                else
                {
                    //未在预期内的结果都返回false
                    return false;
                }
            }
        }
    }
}
