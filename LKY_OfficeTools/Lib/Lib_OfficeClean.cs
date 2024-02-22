/*
 *      [LKY Office Tools] Copyright (C) 2022 - 2024 LiuKaiyuan Inc.
 *      
 *      FileName : Lib_OfficeClean.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static LKY_OfficeTools.Common.Com_InstallerOS;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo.AppPath;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_OfficeClean
    {
        internal static bool RemoveAllOffice()
        {
            try
            {
                //检查是否存在运行中的进程
                var office_p_info = Lib_OfficeProcess.GetRuningProcess();
                if (office_p_info != null && office_p_info.Count > 0)
                {
                    new Log($"\n------> 正在关闭 Office 组件（如果您有未保存的 Office 文档，请立即保存并关闭）...", ConsoleColor.DarkCyan);
                    new Log($"        注意：如果您卡在上述流程达到 1 分钟以上，请您重启计算机，并再次运行本软件！", ConsoleColor.Gray);
                    Lib_OfficeProcess.KillOffice.All();       //友好的结束进程
                    new Log($"     √ 已完成 Office 进程处理。", ConsoleColor.DarkGreen);
                }

                //先使用 ODT 模式卸载，其只能卸载使用 ODT 安装的2016及其以上版本的 Office，但是其耗时短。
                Uninstall.ByODT();

                //再使用 卸载早期版本 模式，可帮助后面 SaRA 模式省时间。
                Uninstall.RemovePreviousVersion();

                //然后使用 SaRA 模式，因为它可以尽可能卸载所有 Office 版本（非ODT），但是耗时长
                Uninstall.BySaRA();

                //无论哪种方式清理，都要再检查一遍是否卸载干净。如果 当前系统 Office 版本数量 > 0，启动强制模式
                var installed_office = OfficeLocalInfo.GetArchiDir();
                var license_list = OfficeLocalInfo.LicenseInfo();
                if (
                    (installed_office != null && installed_office.Count > 0) ||     //存在残留的Office注册表/文件目录
                    license_list != null && license_list.Count > 0                  //存在残留的许可证信息
                    )                                                               //二者满足任意一个时，执行强行清理
                {
                    Uninstall.ForceDelete();    //无论清除是否成功，都继续安装新 office
                }

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }

        internal class Activate
        {
            internal static bool Delete()
            {
                try
                {
                    //获取激活信息
                    var office_installed_key = OfficeLocalInfo.LicenseInfo();

                    if (office_installed_key != null && office_installed_key.Count > 0)
                    {
                        foreach (var now_key in office_installed_key)
                        {
                            string cmd_switch_cd = $"pushd \"{Documents.SDKs.SDKs_Root + @"\Activate"}\"";                  //切换至OSPP文件目录
                            string cmd_remove = $"cscript ospp.vbs /unpkey:{now_key}";
                            string result_log = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_remove})");
                            if (result_log.Contains("success"))
                            {
                                new Log($"     √ 已移除 {now_key} 激活信息。", ConsoleColor.DarkGreen);
                            }
                            else
                            {
                                new Log(result_log);    //获取失败原因
                                new Log($"     × {now_key} 激活信息移除失败！", ConsoleColor.DarkRed);
                                return false;   //有一个产品卸载失败了，就直接返回 false。
                            }
                        }

                        //逐一卸载后，若都为 success，则再执行一次检查
                        office_installed_key = OfficeLocalInfo.LicenseInfo();   //再度获取list
                        if (office_installed_key.Count == 0)   //为0，视为成功
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }

                    return true;    //激活信息为空，或者移除成功时 返回 true
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }
        }

        internal class Uninstall
        {
            internal static bool ForceDelete()
            {
                try
                {
                    new Log($"\n------> 正在执行 Office 强行清理 ...", ConsoleColor.DarkCyan);

                    //移除激活信息
                    Activate.Delete();

                    //删除文件
                    try
                    {
                        //有些文件可能无法彻底删除，但如果后面复查时，不影响安装，则不会返回 false
                        var office_installed_dir = OfficeLocalInfo.GetArchiDir();
                        if (office_installed_dir != null && office_installed_dir.Count > 0)
                        {
                            foreach (var now_dir in office_installed_dir)     //遍历查询所有目录，将其删除
                            {
                                if (Directory.Exists(now_dir))
                                {
                                    Directory.Delete(now_dir, true);
                                }
                            }
                        }
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                    }

                    //清理注册表残余。Office 有些注册表是无法清理干净的，所以用try
                    try
                    {
                        //清理x32注册表
                        Register.DeleteItem(RegistryHive.LocalMachine, RegistryView.Registry32, @"SOFTWARE\Microsoft", "Office");

                        //清理x64注册表。当且仅当，系统是x64系统时，清理x64的注册表节点
                        if (Environment.Is64BitOperatingSystem)
                        {
                            Register.DeleteItem(RegistryHive.LocalMachine, RegistryView.Registry64, @"SOFTWARE\Microsoft", "Office");
                        }
                    }
                    catch {/*注册表 common 项目 铁定删不掉*/}

                    //清除开始菜单
                    try
                    {
                        string root = $@"{Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)}\Programs";
                        var root_dir = Directory.GetDirectories(root, "Microsoft Office*", SearchOption.TopDirectoryOnly);
                        foreach (var now_dir in root_dir)
                        {
                            if (Directory.Exists(now_dir))
                            {
                                Directory.Delete(now_dir, true);
                            }
                        }

                        root = $@"{Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)}\Programs";
                        root_dir = Directory.GetDirectories(root, "Microsoft Office*", SearchOption.TopDirectoryOnly);
                        foreach (var now_dir in root_dir)
                        {
                            if (Directory.Exists(now_dir))
                            {
                                Directory.Delete(now_dir, true);
                            }
                        }
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                    }

                    //复查是否干净了
                    var installed_info = OfficeLocalInfo.GetArchiDir();
                    if (installed_info != null && installed_info.Count > 0)
                    {
                        throw new Exception();
                    }
                    else
                    {
                        new Log($"     √ 已彻底清除 Office 所有组件。", ConsoleColor.DarkGreen);
                        return true;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    new Log($"     × 暂时无法彻底清除 Office 所有组件！", ConsoleColor.DarkRed);
                    return false;
                }
            }

            internal static bool RemovePreviousVersion()
            {
                try
                {
                    new Log($"\n------> 正在卸载 Office 早期版本 ...", ConsoleColor.DarkCyan);

                    //获取 installer 目录
                    string installer_dir = Environment.GetFolderPath(Environment.SpecialFolder.Windows) + "\\Installer";

                    //获取 installer 目录下所有msi文件
                    var msi_files = Directory.GetFiles(installer_dir, "*.msi", SearchOption.TopDirectoryOnly);

                    //生成 产品ID 与 命令行 的字典
                    Dictionary<string, string> msi_id_cmd_dic = new Dictionary<string, string>();
                    //生成 产品ID 与 MSI名称 的字典
                    Dictionary<string, string> msi_id_name_dic = new Dictionary<string, string>();

                    //非空判断
                    if (msi_files == null)
                    {
                        new Log($"      * 未发现 Office 早期版本的维护信息，已跳过此流程。", ConsoleColor.DarkMagenta);
                        return true;
                    }

                    new Log($"     >> 启动 Office 早期版本筛查 ...", ConsoleColor.DarkYellow);

                    //找出符合条件的 Office MSI 文件，并添加至字典
                    foreach (var now_msi in msi_files)
                    {
                        var now_msi_name = GetProductInfo(now_msi, MsiInfoType.ProductName);

                        //非空判断
                        if (string.IsNullOrEmpty(now_msi_name))
                        {
                            new Log($"     × 无法获取 Office 早期版本信息！", ConsoleColor.DarkRed);
                            return false;
                        }

                        if (
                            now_msi_name.Contains("Microsoft Office") &&                        //只查找 Office MSI
                            (now_msi_name.Contains("2003") || now_msi_name.Contains("2007")) && //只查找03、07版本。2010版本（含）以上无法使用本方法卸载
                            !now_msi_name.Contains("MUI") &&                                    //剔除 MUI 语言包
                            !now_msi_name.Contains("Component") &&                              //过滤子组件MSI，否则会导致卸载主程序失败
                            !now_msi_name.Contains("(") &&                                      //过滤其他的语言包，这种MSI通常会用 (English) 这种描述
                            !now_msi_name.Contains("-")                                         //过滤 Microsoft Office Proofing Tools 2013 - English 这种组件
                           )
                        {
                            //获得 MSI 的 ID 号
                            string id = GetProductInfo(now_msi, MsiInfoType.ProductCode);

                            //组成命令行
                            string cmd = $"/uninstall {now_msi} /passive /norestart";

                            //设置 ID/命令行字典
                            msi_id_cmd_dic[id] = cmd;

                            //设置 ID/产品名字典
                            msi_id_name_dic[id] = now_msi_name;
                        }
                    }

                    //非空判断
                    if (msi_id_cmd_dic.Count == 0 || msi_id_name_dic.Count == 0)
                    {
                        new Log($"      * 未发现 Office 早期版本，已跳过此流程。", ConsoleColor.DarkMagenta);
                        return true;
                    }
                    else
                    {
                        new Log($"     √ 已完成 Office 早期版本筛查。", ConsoleColor.DarkGreen);
                    }

                    //逐一卸载 Office
                    foreach (var now_id_cmd in msi_id_cmd_dic)
                    {
                        string product_name = msi_id_name_dic[now_id_cmd.Key];      //完整的产品名
                        new Log($"\n     >> 开始移除 {product_name} 及其组件 ...", ConsoleColor.DarkYellow);

                        //运行卸载
                        var msi_result = Com_ExeOS.Run.Exe("msiexec.exe", now_id_cmd.Value);
                        if (msi_result == -920921)
                        {
                            throw new Exception();
                        }

                        new Log($"        若长时间无 {product_name.Replace("Microsoft Office ", "")} 卸载界面，您可到: 控制面板 -> 程序和功能 列表中手动卸载。", ConsoleColor.Gray);

                        //循环等待结束
                        while (true)
                        {
                            //一直获取对应 产品ID 的注册表
                            var uninstall_reg = Register.Read.AllValues(RegistryHive.LocalMachine,
                                $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{now_id_cmd.Key}", "UninstallString");

                            //一直获取标题包含 Microsoft Office 的进程信息
                            var uninstall_pro_list = Com_ExeOS.Info.GetProcessByTitle("Microsoft Office");
                            if (
                                (uninstall_reg == null || uninstall_reg.Count == 0) &&              //注册表的卸载信息已清空
                                (uninstall_pro_list == null || uninstall_pro_list.Count == 0)       //进程中不存在 Microsoft Office 标题的进程
                                )
                            {
                                break;
                            }

                            Thread.Sleep(3000);     //延迟，防止轮询资源占用过大
                        }

                        new Log($"     √ 已完成 {product_name} 组件移除。", ConsoleColor.DarkGreen);
                    }

                    Thread.Sleep(1000);     //基于体验短暂延迟

                    new Log($"\n     √ 已完成 所有 Office 早期版本卸载。", ConsoleColor.DarkGreen);

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    new Log($"     × 卸载 Office 早期版本出现意外！", ConsoleColor.DarkRed);
                    return false;
                }
            }

            internal static bool BySaRA()
            {
                try
                {
                    new Log($"\n------> 正在卸载 Office 冗余版本 ...", ConsoleColor.DarkCyan);

                    //获取目前存留的版本
                    var install_list = OfficeLocalInfo.GetArchiDir();
                    if (install_list == null || install_list.Count == 0)
                    {
                        //已经不存在残留版本时，直接返回卸载成功
                        new Log($"      * 未发现 Office 冗余版本，已跳过此流程。", ConsoleColor.DarkMagenta);
                        return true;
                    }

                    //定义SaRA文件位置
                    string SaRA_path_root = Documents.SDKs.SDKs_Root + @"\SaRA";
                    string SaRA_path_exe = SaRA_path_root + @"\SaRACmd.exe";

                    //检查SaRA文件是否存在
                    if (!File.Exists(SaRA_path_exe))
                    {
                        new Log($"     × 目录 {SaRA_path_root} 下文件丢失！", ConsoleColor.DarkRed);
                        return false;
                    }

                    //实际测试，平均卸载1个Office需要5分钟的时间，留出1.5倍时间。
                    new Log($"     >> 此过程大约会在 {Math.Ceiling(install_list.Count * 5 * 1.5f)} 分钟内完成，实际用时取决于您的电脑配置，请耐心等待 ...", ConsoleColor.DarkYellow);

                    //执行卸载命令
                    string cmd_switch_cd = $"pushd \"{SaRA_path_root}\"";             //切换至SaRA文件目录
                    string cmd_uninstall = $"SaRACmd.exe -S OfficeScrubScenario -AcceptEula -Officeversion All";
                    string uninstall_result = Com_ExeOS.Run.Cmd($"({cmd_switch_cd})&({cmd_uninstall})");
                    if (!uninstall_result.ToLower().Contains("successful"))
                    {
                        new Log($"SaRA Exception: \n{uninstall_result}");
                        new Log($"     × 卸载 Office 冗余版本失败！", ConsoleColor.DarkRed);
                        return false;
                    }

                    /* 暂时不启用

                    //判断是否需要重启
                    if (uninstall_result.ToLower().Contains("restart"))
                    {
                        //需要重启
                        new Log($"     √ 已完成 Office 卸载操作。为确保卸载干净，请您重启电脑后再次运行本程序。", ConsoleColor.DarkGreen);
                    }
                    else
                    {
                        //不需要重启
                        new Log($"     √ 已完成 Office 卸载操作。", ConsoleColor.DarkGreen);
                    }

                    */

                    new Log($"     √ 已完成 Office 冗余版本卸载。", ConsoleColor.DarkGreen);

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    new Log($"     × 卸载 Office 冗余版本出现异常！", ConsoleColor.DarkRed);
                    return false;
                }
            }

            internal static bool ByODT()
            {
                try
                {
                    new Log($"\n------> 正在卸载 Office ODT 版本 ...", ConsoleColor.DarkCyan);

                    //获取目前存留的 ODT 版本
                    var install_list = Register.Read.AllValues(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");
                    if (install_list == null || install_list.Count == 0)
                    {
                        //已经不存在残留 ODT 版本时，直接返回卸载成功
                        new Log($"      * 未发现 Office ODT 版本，已跳过此流程。", ConsoleColor.DarkMagenta);
                        return true;
                    }

                    //定义ODT文件位置
                    string ODT_path_root = Documents.SDKs.SDKs_Root + @"\ODT";
                    string ODT_path_exe = ODT_path_root + @"\ODT.exe";
                    string ODT_path_xml = ODT_path_root + @"\uninstall.xml";    //此文件需要新生成

                    //生成卸载xml
                    string xml_content = "<Configuration>\n  <Remove All=\"TRUE\" />\n  <Display Level=\"NONE\" AcceptEULA=\"TRUE\"/>\n</Configuration>";
                    File.WriteAllText(ODT_path_xml, xml_content);

                    //检查ODT文件是否存在
                    if (!File.Exists(ODT_path_exe) || !File.Exists(ODT_path_xml))
                    {
                        new Log($"     × 目录 {ODT_path_root} 下文件丢失！", ConsoleColor.DarkRed);
                        return false;
                    }

                    //测试表明，卸载1个ODT版本大约需要2分钟时间，留出2倍的富余
                    new Log($"     >> 此过程大约会在 {Math.Ceiling(install_list.Count * 2 * 2.0f)} 分钟内完成，具体时间取决于您的电脑配置，请稍候 ...", ConsoleColor.DarkYellow);

                    //移除所有激活信息，即使有错误也继续执行后续卸载。
                    if (!Activate.Delete())
                    {
                        new Log("移除激活信息失败！");    //纯打点
                    }

                    new Log($"     >> 卸载仍在继续，请等待其自动完成 ...", ConsoleColor.DarkYellow);
                    //执行卸载命令
                    string uninstall_args = $"/configure \"{ODT_path_xml}\"";
                    var uninstall_code = Com_ExeOS.Run.Exe(ODT_path_exe, uninstall_args);      //卸载

                    var reg_info = Register.Read.AllValues(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                    //未正常结束卸载，视为卸载失败
                    if (uninstall_code == -920921)
                    {
                        new Log($"     × 无法卸载 Office ODT 版本！", ConsoleColor.DarkRed);
                        return false;
                    }

                    //注册表ODT至少存在1个版本时，视为卸载失败
                    if (reg_info != null && reg_info.Count > 0)
                    {
                        //卸载存在问题
                        string err_msg = $"ODT UnInstalling Exception, ExitCode: {uninstall_code}";
                        if (uninstall_code > 0)
                        {
                            //只解析错误码大于0的情况
                            string err_string = string.Empty;
                            if (Lib_OfficeInstall.ODT_Error.TryGetValue(((uint)uninstall_code), out err_string))
                            {
                                err_msg += $"{err_string}";
                            }
                        }

                        new Log(err_msg);         //回调错误码

                        new Log($"     × 已尝试卸载 Office ODT 版本，但系统仍存在 {reg_info.Count} 个无法卸载的版本！", ConsoleColor.DarkRed);
                        return false;
                    }

                    //卸载正常 + 注册表已清空时，开始安装新版本
                    new Log($"     √ 已完成 Office ODT 版本卸载。", ConsoleColor.DarkGreen);
                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    new Log($"     × 卸载 Office ODT 版本出现异常！", ConsoleColor.DarkRed);
                    return false;
                }
            }
        }
    }
}
