/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2023 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInstall.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static LKY_OfficeTools.Common.Com_ExeOS;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppCommand;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppReport;
using static LKY_OfficeTools.Lib.Lib_AppState;
using static LKY_OfficeTools.Lib.Lib_OfficeClean;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo.OfficeLocalInfo;

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

            try
            {
                //版本获取成功时，执行后续操作。
                if (OfficeNetInfo.OfficeLatestVersion != null)
                {
                    //判断是否已经安装了当前版本
                    InstallState install_state = GetOfficeState();
                    if (install_state == InstallState.Correct)                //已安装最新版，无需下载
                    {
                        new Log($"\n      * 当前系统安装了最新 Office 版本，已跳过下载、安装流程。", ConsoleColor.DarkMagenta);

                        //开始激活
                        Lib_OfficeActivate.Activating();
                    }

                    //当不存在 VersionToReport or 其版本与最新版不一致 or 产品ID不一致 or 安装位数与系统不一致时，需要下载新文件。
                    else
                    {
                        //被动模式下，如果版本不正确，不会重新安装，因为用户可能已经卸载。
                        if (Current_RunMode == RunMode.Passive)
                        {
                            new Log($"\n     × 当前系统未安装最新版本的 Office，激活停止！", ConsoleColor.DarkRed);
                            return;
                        }

                        //如果是主动模式，则下载并安装最新版
                        else if (Current_RunMode == RunMode.Manual)
                        {
                            //开始下载，并获得下载结果
                            int DownCode = Lib_OfficeDownload.StartDownload();

                            //判断下载情况
                            switch (DownCode)
                            {
                                //下载成功，开始安装
                                case 1:
                                    //冲突检查
                                    if (ConflictCheck())
                                    {
                                        //通过进入安装，不通过直接返回false
                                        if (StartInstall())
                                        {
                                            //安装成功，进入激活程序
                                            Lib_OfficeActivate.Activating();
                                        }
                                        else
                                        {
                                            new Log($"\n     × 因 Office 安装失败，已跳过激活流程！", ConsoleColor.DarkRed);
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        new Log($"\n     × 无法完成 Office 冲突性检查，已跳过安装流程！", ConsoleColor.DarkRed);
                                        break;
                                    }
                                    break;
                                case 0:
                                    new Log($"\n     × 未能找到可用的 Office 安装文件，已跳过安装流程！", ConsoleColor.DarkRed);
                                    break;
                                case -1:
                                    //用户中止了下载
                                    new Log($"\n     × 用户取消了下载 Office 安装文件，已跳过安装流程！", ConsoleColor.DarkRed);
                                    return;
                            }
                        }

                        //其他模式跳出
                        else
                        {
                            new Log($"{Current_RunMode} 模式下，暂不支持下载、安装、激活操作。");
                            return;
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                new Log($"\n     × Office 安装过程中发生异常，已跳过安装流程！", ConsoleColor.DarkRed);
                new Log(Ex.ToString());
            }
            finally
            {
                //全部完成后，判断是否成功
                if (Current_StageType != ProcessStage.Finish_Success)
                {
                    //只要全部流程结束后，不是成功状态（并且没有中断情况），就设置为 失败 
                    Current_StageType = ProcessStage.Finish_Fail;
                }
            }
        }

        /// <summary>
        /// 冲突版本检查，卸载和主发行版本不一样的的 Office。
        /// 无冲突 或 冲突已经解决返回 true，有冲突且无法解决返回 false。
        /// </summary>
        internal static bool ConflictCheck()
        {
            try
            {
                new Log($"\n------> 正在检查 Office 安装环境 ...", ConsoleColor.DarkCyan);

                //--------------------------------------------- 先检查注册表 ---------------------------------------------
                bool regdir_pass = false;      //默认是冲突的

                var Current_Office_Dir = GetArchiDir();         //不能用 Full 函数判断是否 null，本字典始终有 key 的，所以恒不为 null
                if (Current_Office_Dir != null && Current_Office_Dir.Count > 0)     //非空且为大于0时
                {
                    //注册表只有1个值，接下来判断这个值属于哪个架构
                    if (Current_Office_Dir.Count == 1)
                    {
                        //判断仅有的1个office架构是不是odt架构（odt下面的 x32 x64 有1个不为空，则再判断是否符合 位数 和 发行ID）
                        string odt_x32_reg = GetArchiDirFull()[OfficeArchi.Office_ODT_x32];
                        string odt_x64_reg = GetArchiDirFull()[OfficeArchi.Office_ODT_x64];
                        if (!string.IsNullOrEmpty(odt_x32_reg) || !string.IsNullOrEmpty(odt_x64_reg))
                        {
                            //64位系统额外增加位数判断
                            if (Environment.Is64BitOperatingSystem)
                            {
                                if (!string.IsNullOrEmpty(odt_x32_reg))
                                {
                                    //x64系统，但是在注册表x32路径有值，冲突
                                    regdir_pass = false;
                                }
                                else
                                {
                                    //x64系统，虽然在注册表x64路径有值了，但是还需判断是否 平台信息 一致
                                    string platform = Register.Read.Value(RegistryHive.LocalMachine, RegistryView.Registry64, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "Platform");
                                    if (string.IsNullOrWhiteSpace(platform) || platform == "x86")
                                    {
                                        regdir_pass = false;       //虽然出现在了与x64系统注册表匹配的路径，但是Office平台版本并非x64
                                    }
                                }
                            }

                            //odt位数与系统匹配，开始校验产品ID
                            ///获取其ID是否和当前的ID一致，不一致，则false。根据系统位数来取，如果用户是x64系统，则取的是x64的software注册表。x32同理
                            string Current_Office_ID = Register.Read.ValueBySystem(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                            ///目标ID
                            string Pop_Office_ID = "2021Volume";

                            //ODT 只安装了1个版本，但 Current_Office_ID 为空，视为冲突
                            if (string.IsNullOrEmpty(Current_Office_ID))
                            {
                                regdir_pass = false;
                            }
                            else
                            {
                                //替换支持的版本ID。如果用户还安装了其它架构，进行替换后，该值不是空
                                var diff_products = Current_Office_ID
                                    .Replace($"ProjectPro{Pop_Office_ID}", "")
                                    .Replace($"ProPlus{Pop_Office_ID}", "")
                                    .Replace($"VisioPro{Pop_Office_ID}", "")
                                    .Replace(",", "");

                                //判断有没有额外的版本ID
                                if (string.IsNullOrWhiteSpace(Current_Office_ID))
                                {
                                    //为空，无额外的ID，是不冲突的
                                    regdir_pass = true;
                                }
                                else
                                {
                                    //不为空，说明还安装了其他的版本，如：ProPlus2016Volume，则冲突
                                    regdir_pass = false;
                                }
                            }
                        }
                        else
                        {
                            //不是odt架构，直接false
                            regdir_pass = false;
                        }
                    }
                    else
                    {
                        //如果注册表值大于1个，也就是安装了多个office，直接设置为冲突，设为false
                        //如果用户的注册表存在两个ODT同版本不同位的情况（x32、x64都有），则也视为有冲突。
                        regdir_pass = false;
                    }
                }
                else
                {
                    //注册表没有值，或者 有值但目录不存在，无冲突，true
                    regdir_pass = true;
                }
                //--------------------------------------------- 检查注册表 结束 ---------------------------------------------



                //--------------------------------------------- 再检查许可证 ---------------------------------------------
                bool license_pass = false;     //默认是冲突的

                var installed_key = LicenseInfo();       //获取许可证列表
                if (installed_key != null && installed_key.Count > 0)
                {
                    //注册表有值

                    //获取当前的许可证信息
                    string cmd_switch_cd = $"pushd \"{AppPath.Documents.SDKs.Activate}\"";                  //切换至OSPP文件目录
                    string cmd_installed_info = "cscript ospp.vbs /dstatus";                                //查看激活状态
                    string installed_license_info = Run.Cmd($"({cmd_switch_cd})&({cmd_installed_info})");     //查看所有版本激活情况

                    //判断值的数量
                    if (installed_key.Count == 1)
                    {
                        //判断安装的许可证是否是目标大版本
                        if (installed_license_info.Contains("ProPlus2021VL"))
                        {
                            //是目标大版本
                            license_pass = true;
                        }
                        else
                        {
                            //不是本程序的大版本，直接false
                            license_pass = false;
                        }
                    }
                    else if (installed_key.Count == 2)
                    {
                        //检测到的2个版本是否是支持的3个大版本其中的2个版本。共有3种组合
                        if (
                            (installed_license_info.Contains("ProPlus2021VL") && installed_license_info.Contains("ProjectPro2021VL")) |
                            (installed_license_info.Contains("ProPlus2021VL") && installed_license_info.Contains("VisioPro2021VL")) |
                            (installed_license_info.Contains("ProjectPro2021VL") && installed_license_info.Contains("VisioPro2021VL"))
                            )
                        {
                            //是目标大版本
                            license_pass = true;
                        }
                        else
                        {
                            //不是本程序的大版本，直接false
                            license_pass = false;
                        }
                    }
                    else if (installed_key.Count == 3)
                    {
                        //安装3个许可证，必须是指定的三个版本，否则就是有别的许可证，属于冲突
                        if (installed_license_info.Contains("ProPlus2021VL")
                            & installed_license_info.Contains("ProjectPro2021VL")
                            & installed_license_info.Contains("VisioPro2021VL"))
                        {
                            //是目标大版本
                            license_pass = true;
                        }
                        else
                        {
                            //不是本程序的大版本，直接false
                            license_pass = false;
                        }
                    }
                    else
                    {
                        //如果大于3个，也就是安装了多个office，直接设置为冲突，设为false
                        license_pass = false;
                    }
                }
                else
                {
                    //许可证没有值，无冲突，true
                    license_pass = true;
                }
                //--------------------------------------------- 检查许可证 结束 ---------------------------------------------


                //聚合判断是否有冲突，必须二者同时通过，才是不冲突。只要注册表或者许可证，有一个出现了false，就视为冲突
                if (regdir_pass && license_pass)
                {
                    //不存在版本冲突时，直接开始安装
                    new Log($"     √ 已通过 Office 安装环境检查。", ConsoleColor.DarkGreen);
                    return true;
                }
                else
                {
                    new Log($"     ★ 发现 {Current_Office_Dir.Count + installed_key.Count} 个冲突的 Office 版本，若要继续，必须先卸载旧版。", ConsoleColor.Gray);

                    //判断是否包含自动卸载标记
                    if (!AppCommandFlag.HasFlag(ArgsFlag.Auto_Remove_Conflict_Office))
                    {
                        if (KeyMsg.Confirm("确认安装新版 Office 并卸载旧版本及其插件"))
                        {
                            new Log($"     √ 您已主动确认 卸载 Office 所有旧版本。", ConsoleColor.DarkGreen);
                            Thread.Sleep(500);     //基于体验，稍微停留下

                            //获取卸载结果
                            var result = RemoveAllOffice();

                            //卸载后提示重启电脑
                            if (KeyMsg.Choose("建议您在安装新版本 Office 前，重启计算机，以确保旧版 Office 卸载干净！"))
                            {
                                //确认重启
                                new Log($"     √ 您已主动确认 重启计算机。系统将在 1分钟 内重启，请注意保存文件。", ConsoleColor.DarkGreen);

                                //执行重启命令行（shutdown.exe -r -t 3600）
                                Run.Cmd("shutdown.exe -r -t 60");

                                //重启打点&退出
                                Pointing(ProcessStage.RestartPC, true);
                                KeyMsg.Quit(128);
                            }
                            else
                            {
                                //跳过了重启选择
                                new Log($"      * 您已拒绝 重启计算机，若安装失败，您可在重启后重新运行本工具。", ConsoleColor.DarkMagenta);
                                Thread.Sleep(500);    //基于体验，稍微停留下
                            }

                            return result;
                        }
                        else
                        {
                            new Log($"     × 您已拒绝 卸载 Office 其他版本，新版本无法安装！", ConsoleColor.DarkRed);
                            return false;
                        }
                    }
                    else
                    {
                        //有自动卸载的标记，直接开始卸载
                        new Log($"      * 自动卸载冲突 Office 模式下，自动确认并开始 ...", ConsoleColor.DarkMagenta);
                        return RemoveAllOffice();
                    }
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }

        /// <summary>
        /// 开始安装 Office
        /// </summary>
        internal static bool StartInstall()
        {
            try
            {
                //定义ODT文件位置
                string ODT_path_root = AppPath.Documents.SDKs.SDKs_Root + @"\ODT";
                string ODT_path_exe = ODT_path_root + @"\ODT.exe";
                string ODT_path_xml = ODT_path_root + @"\config.xml";

                //检查ODT文件是否存在
                if (!File.Exists(ODT_path_exe) || !File.Exists(ODT_path_xml))
                {
                    new Log($"     × 目录 {ODT_path_root} 下文件丢失！", ConsoleColor.DarkRed);
                    return false;
                }

                //修改新的xml信息
                ///修改安装目录，安装目录为运行根目录
                bool isNewInstallPath = Com_FileOS.XML.SetValue(ODT_path_xml, "SourcePath", AppPath.ExecuteDir);

                //检查是否修改成功（安装目录）
                if (!isNewInstallPath)
                {
                    new Log($"     × 配置 Install 信息错误！", ConsoleColor.DarkRed);
                    return false;
                }

                ///修改为新版本号
                bool isNewVersion = Com_FileOS.XML.SetValue(ODT_path_xml, "Version", OfficeNetInfo.OfficeLatestVersion.ToString());

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

                ///提示用户要安装的Office组件
                var msg_tip = "\n     ★ 本工具默认安装 Word、PPT、Excel，并可选配更多组件：";
                var msg_index_first = "\n        \tOutlook = 1\tOneNote = 2\tAccess = 3";
                var msg_index_second = "\n        \tVisio = 4\tProject = 5\tPublisher = 6";
                var msg_index_third = "\n        \tTeams = 7\tOneDrive = 8\tLync = 9";
                var msg_input = "\n        如安装：Outlook、OneNote、Visio、，请输入：1,2,4 后回车，如不增加组件，请直接按回车键。";
                new Log(msg_tip + msg_index_first + msg_index_second + msg_index_third + msg_input, ConsoleColor.Gray);

                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write("\n        请输入追加的组件序号：");

                //去除非法字符
                var add_install = Console.ReadLine().Trim(' ').Trim(',')
                    .Replace("，", ",").Replace(",,", ",").Replace("，，", ",")
                    .Replace(".", ",").Replace("。", ",");

                //读取配置全部内容
                var config = File.ReadAllText(ODT_path_xml);

                //默认不安装Visio、Project
                bool install_visio = false;
                bool install_project = false;

                //非空判断
                if (!string.IsNullOrWhiteSpace(add_install))
                {
                    //安装附加组件
                    var add_install_list = add_install.Split(',');
                    //检查输入的序号是否正确
                    if (add_install_list.Length == 0)
                    {
                        //只增加1个附加组件时
                        //非法序号区间判断
                        if (add_install.Length > 1 | int.Parse(add_install) > 9 | int.Parse(add_install) < 1)
                        {
                            new Log($"      * 您输入的“{add_install}”序号并非有效组件序号，故工具将默认安装：Word、PPT、Excel 三件套。", ConsoleColor.DarkMagenta);
                            return false;
                        }

                        //判断组件类型
                        switch (int.Parse(add_install))
                        {
                            case 1:
                                //Outlook
                                config = config.Replace("<ExcludeApp ID=\"Outlook\" />", "");
                                break;
                            case 2:
                                //OneNote
                                config = config.Replace("<ExcludeApp ID=\"OneNote\" />", "");
                                break;
                            case 3:
                                //Access
                                config = config.Replace("<ExcludeApp ID=\"Access\" />", "");
                                break;
                            case 4:
                                //Visio
                                install_visio = true;
                                break;
                            case 5:
                                //Project
                                install_project = true;
                                break;
                            case 6:
                                //Publisher
                                config = config.Replace("<ExcludeApp ID=\"Publisher\" />", "");
                                break;
                            case 7:
                                //Teams
                                config = config.Replace("<ExcludeApp ID=\"Teams\" />", "");
                                break;
                            case 8:
                                //OneDrive
                                config = config.Replace("<ExcludeApp ID=\"OneDrive\" />", "");
                                break;
                            case 9:
                                //Lync
                                config = config.Replace("<ExcludeApp ID=\"Lync\" />", "");
                                break;
                        }
                    }
                    else
                    {
                        //遍历要安装的组件
                        foreach (var now_add in add_install_list)
                        {
                            //非法序号区间判断
                            if (now_add.Length > 1 | int.Parse(now_add) > 9 | int.Parse(now_add) < 1)
                            {
                                new Log($"      * 您输入的“{now_add}”序号并非有效组件序号，故工具将默认安装：Word、PPT、Excel 三件套。", ConsoleColor.DarkMagenta);
                                return false;
                            }

                            //判断组件类型
                            switch (int.Parse(now_add))
                            {
                                case 1:
                                    //Outlook
                                    config = config.Replace("<ExcludeApp ID=\"Outlook\" />", "");
                                    break;
                                case 2:
                                    //OneNote
                                    config = config.Replace("<ExcludeApp ID=\"OneNote\" />", "");
                                    break;
                                case 3:
                                    //Access
                                    config = config.Replace("<ExcludeApp ID=\"Access\" />", "");
                                    break;
                                case 4:
                                    //Visio
                                    install_visio = true;
                                    break;
                                case 5:
                                    //Project
                                    install_project = true;
                                    break;
                                case 6:
                                    //Publisher
                                    config = config.Replace("<ExcludeApp ID=\"Publisher\" />", "");
                                    break;
                                case 7:
                                    //Teams
                                    config = config.Replace("<ExcludeApp ID=\"Teams\" />", "");
                                    break;
                                case 8:
                                    //OneDrive
                                    config = config.Replace("<ExcludeApp ID=\"OneDrive\" />", "");
                                    break;
                                case 9:
                                    //Lync
                                    config = config.Replace("<ExcludeApp ID=\"Lync\" />", "");
                                    break;
                            }
                        }

                        new Log($"     √ 检查完毕，本工具将追加安装 {add_install} 组件。", ConsoleColor.DarkGreen);
                    }
                }
                //不安装附加组件的情况
                else { }

                //不安装Viso时，移除相关配置
                if (!install_visio)
                {
                    var remove_info = Com_TextOS.GetCenterText(config, "<Product ID=\"Visio", "</Product>");
                    remove_info = "<Product ID=\"Visio" + remove_info + "</Product>";
                    config = config.Replace(remove_info, "");
                }

                //不安装Project时，移除相关配置
                if (!install_project)
                {
                    var remove_info = Com_TextOS.GetCenterText(config, "<Product ID=\"Project", "</Product>");
                    remove_info = "<Product ID=\"Project" + remove_info + "</Product>";
                    config = config.Replace(remove_info, "");
                }

                //保存修改后的配置
                File.WriteAllText(ODT_path_xml, config);

                //开始安装
                new Log($"\n------> 正在安装 Office v{OfficeNetInfo.OfficeLatestVersion} ...", ConsoleColor.DarkCyan);

                ///先结束掉可能还在安装的 Office 进程（强制结束，不等待）
                Lib_AppSdk.KillAllSdkProcess(KillExe.KillMode.Only_Force);

                ///命令安装
                string install_args = $"/configure \"{ODT_path_xml}\"";     //配置命令行
                var install_code = Run.Exe(ODT_path_exe, install_args);

                //检查是否因配置不正确等导致，意外退出安装
                if (install_code == -920921)
                {
                    new Log($"     × Office v{OfficeNetInfo.OfficeLatestVersion} 安装意外结束！", ConsoleColor.DarkRed);
                    return false;
                }

                //无论是否成功，都增加一步结束进程
                KillExe.ByExeName("OfficeClickToRun", KillExe.KillMode.Only_Force, true);      //结束无关进程
                KillExe.ByExeName("OfficeC2RClient", KillExe.KillMode.Only_Force, true);       //结束无关进程

                //检查安装是否成功
                InstallState install_state = GetOfficeState();

                //安装了最新版
                if (install_state == InstallState.Correct)
                {
                    //安装成功
                    new Log($"     √ 已完成 Office v{OfficeNetInfo.OfficeLatestVersion} 安装。", ConsoleColor.DarkGreen);
                    return true;
                }
                else
                {
                    //安装存在问题
                    string err_msg = $"ODT Installing Exception, ExitCode: {install_code}";
                    if (install_code > 0)
                    {
                        //只解析错误码大于0的情况
                        string err_string = string.Empty;
                        if (ODT_Error.TryGetValue(((uint)install_code), out err_string))
                        {
                            err_msg += $"{err_string}";
                        }
                    }

                    new Log(err_msg);         //回调错误码

                    //未安装
                    if (install_state == InstallState.None)
                    {
                        new Log($"     × 安装失败，未在当前系统检测到任何 Office 版本！", ConsoleColor.DarkRed);
                        new Log(install_state);     //打点失败注册表记录
                        return false;
                    }

                    //包含不同版本
                    if (install_state == InstallState.Diff)
                    {
                        new Log($"     × 已安装的 Office 版本与预期的 v{OfficeNetInfo.OfficeLatestVersion} 版本不符！", ConsoleColor.DarkRed);
                        new Log(install_state);     //打点失败注册表记录
                        return false;
                    }

                    //包含多个版本
                    if (install_state == InstallState.Multi)
                    {
                        //系统存在多个版本
                        new Log($"     × 安装异常，当前系统存在多个 Office 版本！", ConsoleColor.DarkRed);
                        new Log(install_state);     //打点失败注册表记录
                        return false;
                    }

                    //其它未可知情况，视为失败
                    return false;
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }

        /// <summary>
        /// ODT工具常见错误码。
        /// https://learn.microsoft.com/zh-cn/windows/client-management/mdm/office-csp
        /// </summary>
        internal static Dictionary<uint, string> ODT_Error
        {
            get
            {
                Dictionary<uint, string> res = new Dictionary<uint, string>
                {
                    [0] = "安装成功。",
                    [997] = "安装正在进行中。",
                    [13] = "无法验证下载的 Office 部署工具 (ODT) 的签名。",
                    [1460] = "下载 ODT 超时。",
                    [1602] = "用户已取消运行。",
                    [1603] = "未通过任何预检检查。安装 2016 MSI 时尝试安装 SxS。" +
                    "当前安装的 Office 与尝试安装的 Office 之间的位不匹配 (例如，在当前安装 64 位版本时尝试安装 32 位版本时。)",
                    [17000] = "未能启动 C2RClient。",
                    [17001] = "未能在 C2RClient 中排队安装方案。",
                    [17002] = "未能完成该过程。可能的原因：" +
                    "（1）用户已取消安装" +
                    "（2）安装已由另一个安装取消" +
                    "（3）安装期间磁盘空间不足" +
                    "（4）未知语言 ID",
                    [17003] = "另一个方案正在运行。",
                    [17004] = "无法完成需要的清理。可能的原因：" +
                    "（1）未知 SKU" +
                    "（2）CDN 上不存在内容。例如，尝试安装不受支持的 LAP，例如 zh-sg" +
                    "（3）内容不可用的 CDN 问题" +
                    "（4）签名检查问题，例如 Office 内容的签名检查失败" +
                    "（5）用户已取消",
                    [17005] = "ERROR！SCENARIO CANCELLED AS PLANNED。",
                    [17006] = "通过运行应用阻止更新。",
                    [17007] = "客户端在“删除安装”方案中请求客户端清理。",
                    [17100] = "C2RClient 命令行错误。",
                    [0x80004005] = "ODT 不能用于安装批量许可证。",
                    [0x8000ffff] = "尝试在计算机上没有 C2R Office 时卸载。"
                };
                return res;
            }
        }
    }
}
