/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInstall.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using Microsoft.Win32;
using System;
using System.IO;
using System.Threading;
using static LKY_OfficeTools.Common.Com_ExeOS;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppCommand;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
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
                case 2:
                    //已安装最新版，无需下载安装，直接进入激活模块
                    new Lib_OfficeActivate();
                    break;
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
                    //用户中止了下载，不执行附加内容
                    return;
            }

            //全部完成后，判断是否成功
            if (Lib_AppState.Current_StageType != Lib_AppState.ProcessStage.Finish_Success)
            {
                //只要全部流程结束后，不是成功状态（并且没有中断情况），就设置为 失败 
                Lib_AppState.Current_StageType = Lib_AppState.ProcessStage.Finish_Fail;
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

                            ///计划安装的ID
                            string Pop_Office_ID = Com_TextOS.GetCenterText(AppJson.Info, "\"Pop_Office_ID\": \"", "\"");                       //安装ID信息

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

                    //判断值的数量
                    if (installed_key.Count == 1)
                    {
                        string Pop_Office_LicenseName = Com_TextOS.GetCenterText(AppJson.Info, "\"Pop_Office_LicenseName\": \"", "\"");     //许可证信息
                        ///获取失败时，默认使用 ProPlus2021VL 版
                        if (string.IsNullOrEmpty(Pop_Office_LicenseName))
                        {
                            Pop_Office_LicenseName = "ProPlus2021VL";
                        }

                        //判断仅有的1个office是不是本程序即将安装的大版本
                        string cmd_switch_cd = $"pushd \"{AppPath.Documents.SDKs.Activate}\"";                  //切换至OSPP文件目录
                        string cmd_installed_info = "cscript ospp.vbs /dstatus";                                //查看激活状态
                        string installed_license_info = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_installed_info})");     //查看所有版本激活情况

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
                    new Log($"     ★ 发现 {Current_Office_Dir.Count + installed_key.Count} 个冲突的 Office 版本，若要继续，必须先卸载旧版。", ConsoleColor.Gray);

                    //判断是否包含自动卸载标记
                    if (!AppCommandFlag.HasFlag(ArgsFlag.Auto_Remove_Conflict_Office))
                    {
                        if (Lib_AppMessage.KeyMsg.Confirm("确认安装新版 Word、PPT、Excel、Outlook、OneNote、Access 六件套，并卸载旧版本及其插件"))
                        {
                            new Log($"     √ 您已主动确认 卸载 Office 所有旧版本。", ConsoleColor.DarkGreen);
                            Thread.Sleep(500);     //基于体验，稍微停留下
                            return RemoveAllOffice();
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
            string Pop_Office_ID = Com_TextOS.GetCenterText(AppJson.Info, "\"Pop_Office_ID\": \"", "\"");
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
            new Log($"\n------> 正在安装 Office v{OfficeNetVersion.latest_version} ...", ConsoleColor.DarkCyan);

            ///先结束掉可能还在安装的 Office 进程（强制结束，不等待）
            Lib_AppSdk.KillAllSdkProcess(KillExe.KillMode.Only_Force);

            ///命令安装
            string install_args = $"/configure \"{ODT_path_xml}\"";     //配置命令行
            bool isInstallFinish = Run.Exe(ODT_path_exe, install_args);

            //检查是否因配置不正确等导致，意外退出安装
            if (!isInstallFinish)
            {
                new Log($"     × Office v{OfficeNetVersion.latest_version} 安装意外结束！", ConsoleColor.DarkRed);
                return false;
            }

            //无论是否成功，都增加一步结束进程
            KillExe.ByExeName("OfficeClickToRun", KillExe.KillMode.Only_Force, true);      //结束无关进程
            KillExe.ByExeName("OfficeC2RClient", KillExe.KillMode.Only_Force, true);       //结束无关进程

            //检查安装是否成功
            InstallState install_state = GetOfficeState();

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
                new Log($"     × 已安装的 Office 版本与预期的 v{OfficeNetVersion.latest_version} 版本不符！", ConsoleColor.DarkRed);
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

            //安装了最新版
            if (install_state == InstallState.Correct)
            {
                //安装成功
                new Log($"     √ 已完成 Office v{OfficeNetVersion.latest_version} 安装。", ConsoleColor.DarkGreen);
                return true;
            }

            //其它未可知情况，视为失败
            return false;
        }
    }
}
