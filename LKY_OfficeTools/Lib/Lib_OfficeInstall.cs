/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInstall.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.App;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_OfficeClean;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo.OfficeLocalInstall;

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
                    if (ConflictCheck())        //冲突检查
                    {
                        //通过进入安装，不通过直接返回false
                        if (StartInstall())
                        {
                            //安装成功，进入激活程序
                            new Lib_OfficeActivate();
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
                    //已安装最新版，无需下载安装，直接进入激活模块
                    new Lib_OfficeActivate();
                    break;
            }

            //全部完成后，判断是否成功
            if (State.Current_Runtype != State.RunType.Finish_Success)
            {
                //只要全部流程结束后，不是成功状态，就设置为 失败 
                State.Current_Runtype = State.RunType.Finish_Fail;
            }
        }

        /// <summary>
        /// 冲突版本检查，卸载和主发行版本不一样的的 Office。
        /// 有冲突，返回false，无冲突，返回true。
        /// </summary>
        internal static bool ConflictCheck()
        {
            try
            {
                new Log($"\n------> 正在检查 Office 安装环境 ...", ConsoleColor.DarkCyan);

                //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                string info = Lib_AppUpdate.latest_info;
                if (string.IsNullOrEmpty(info))
                {
                    info = Com_WebOS.Visit_WebClient(Lib_AppUpdate.update_json_url);
                }

                //--------------------------------------------- 先检查注册表 ---------------------------------------------
                bool regdir_pass = false;      //默认是冲突的

                var Current_Office_Dir = GetArchiDir();         //不能用 Full 函数判断是否 null，本字典始终有 key 的，所以恒不为 null
                if (Current_Office_Dir != null && Current_Office_Dir.Count > 0)     //非空且为大于0时
                {
                    //注册表有值，且目录真实存在
                    if (Current_Office_Dir.Count == 1)
                    {
                        //判断仅有的1个office架构是不是odt架构
                        if (GetArchiDirFull()[OfficeArchi.Office_ODT] != null)
                        {
                            //是odt架构
                            ///获取其ID是否和当前的ID一致，不一致，则false
                            string Current_Office_ID = Register.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                            ///计划安装的ID
                            string Pop_Office_ID = Com_TextOS.GetCenterText(info, "\"Pop_Office_ID\": \"", "\"");                       //安装ID信息

                            ///获取失败时，默认使用 ProPlus2021Volume 版
                            if (string.IsNullOrEmpty(Pop_Office_ID))
                            {
                                Pop_Office_ID = "ProPlus2021Volume";
                            }

                            //判断ID是否相同
                            if (!string.IsNullOrEmpty(Current_Office_ID) && Current_Office_ID == Pop_Office_ID)
                            {
                                //ID相等，是不冲突的
                                regdir_pass = true;
                            }
                            else
                            {
                                //不相等，则冲突
                                regdir_pass = false;
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

                    //判断值的数量
                    if (installed_key.Count == 1)
                    {
                        string Pop_Office_LicenseName = Com_TextOS.GetCenterText(info, "\"Pop_Office_LicenseName\": \"", "\"");     //许可证信息
                        ///获取失败时，默认使用 ProPlus2021VL 版
                        if (string.IsNullOrEmpty(Pop_Office_LicenseName))
                        {
                            Pop_Office_LicenseName = "ProPlus2021VL";
                        }

                        //判断仅有的1个office是不是本程序即将安装的大版本
                        string cmd_switch_cd = $"pushd \"{App.Path.SDK.OSPP_Dir}\"";                  //切换至OSPP文件目录
                        string cmd_installed_info = "cscript ospp.vbs /dstatus";                                //查看激活状态
                        string installed_license_info = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_installed_info})");     //查看所有版本激活情况

                        //判断安装的许可证是否是目标大版本
                        if (installed_license_info.Contains(Pop_Office_LicenseName))
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
                        //如果大于1个，也就是安装了多个office，直接设置为冲突，设为false
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
                    new Log($"     ☆ 发现冲突的 Office 版本，如需安装最新版，必须先卸载旧版。", ConsoleColor.Gray);

                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.Write($"        确认安装新版 Word、PPT、Excel、Outlook、OneNote、Access 六件套，并卸载旧版本及其插件，请按 回车键 继续 ...");
                    if (Console.ReadKey().Key == ConsoleKey.Enter)
                    {
                        new Log($"\n     √ 您已主动确认 卸载 Office 所有旧版本。", ConsoleColor.DarkGreen);

                        //先使用 ODT 模式卸载，其只能卸载使用 ODT 安装的2016及其以上版本的 Office，但是其耗时短。
                        Uninstall.ByODT();
                        new Log($"\n     >> 第二阶段卸载正在进行，请稍候 ...", ConsoleColor.DarkYellow);
                        //第二阶段使用 SaRA 模式，因为它可以尽可能卸载所有 Office 版本（非ODT），但是耗时长
                        Uninstall.BySaRA();

                        //无论哪种方式清理，都要再检查一遍是否卸载干净。如果 当前系统 Office 版本数量 > 0，启动强制模式
                        var installed_office = GetArchiDir();
                        if (installed_office != null && installed_office.Count > 0)
                        {
                            Uninstall.ForceDelete();    //无论清除是否成功，都继续安装新 office
                        }

                        return true;
                    }
                    else
                    {
                        new Log($"\n     × 您已拒绝 卸载 Office 其他版本，新版本无法安装！", ConsoleColor.DarkRed);
                        return false;
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
            //定义ODT文件位置
            string ODT_path_root = App.Path.SDK.Root + @"\ODT";
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
            string info = Lib_AppUpdate.latest_info;
            if (string.IsNullOrEmpty(info))
            {
                info = Com_WebOS.Visit_WebClient(Lib_AppUpdate.update_json_url);
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
            new Log($"\n------> 开始安装 Office v{OfficeNetVersion.latest_version} ...", ConsoleColor.DarkCyan);

            ///先结束掉可能还在安装的 Office 进程
            Com_ProcessOS.KillProcess("OfficeClickToRun");
            Com_ProcessOS.KillProcess("OfficeC2RClient");
            Com_ProcessOS.KillProcess("ODT");

            ///命令安装
            string install_args = $"/configure \"{ODT_path_xml}\"";     //配置命令行
            bool isInstallFinish = Com_ExeOS.RunExe(ODT_path_exe, install_args);

            //检查是否因配置不正确等导致，意外退出安装
            if (!isInstallFinish)
            {
                new Log($"     × Office v{OfficeNetVersion.latest_version} 安装意外结束！", ConsoleColor.DarkRed);
                return false;
            }

            //无论是否成功，都增加一步结束进程
            Com_ProcessOS.KillProcess("OfficeClickToRun");      //结束无关进程
            Com_ProcessOS.KillProcess("OfficeC2RClient");       //结束无关进程

            //检查安装是否成功
            InstallState install_state = GetOfficeState();

            //未安装
            if (install_state.HasFlag(InstallState.None))
            {
                new Log($"     × 安装失败，未在当前系统检测到任何 Office 版本！", ConsoleColor.DarkRed);
                new Log(install_state);     //打点失败注册表记录
                return false;
            }

            //包含不同版本
            if (install_state.HasFlag(InstallState.Diff))
            {
                new Log($"     × 已安装的 Office 版本与预期的 v{OfficeNetVersion.latest_version} 版本不符！", ConsoleColor.DarkRed);
                new Log(install_state);     //打点失败注册表记录
                return false;
            }

            //安装了最新版
            if (install_state.HasFlag(InstallState.Correct))
            {
                //包含多个版本
                if (install_state.HasFlag(InstallState.Multi))
                {
                    //虽然安装成功，但是还有别的版本
                    new Log($"     √ Office v{OfficeNetVersion.latest_version} 已安装完成。", ConsoleColor.DarkGreen);
                    new Log($"     ☆ 但系统存在多个 Office 版本，若 Office 激活失败，请您卸载其它版本后，重新运行本软件。", ConsoleColor.Gray);
                    new Log(install_state);     //打点失败注册表记录
                    return true;
                }
                else
                {
                    //独占性安装成功
                    new Log($"     √ Office v{OfficeNetVersion.latest_version} 已安装完成。", ConsoleColor.DarkGreen);
                    return true;
                }
            }

            //其它未可知情况，视为失败
            return false;
        }
    }
}
