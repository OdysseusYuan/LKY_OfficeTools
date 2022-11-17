/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeClean.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_OfficeInfo;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 清理 Office 的类库
    /// </summary>
    internal class Lib_OfficeClean
    {
        /// <summary>
        /// 解除激活信息
        /// </summary>
        internal class Activate
        {
            /// <summary>
            /// 移除所有激活信息
            /// </summary>
            /// <returns></returns>
            internal static bool RemoveAll()
            {
                try
                {
                    //获取激活信息
                    var office_installed_key = OfficeLocalInstall.LicenseInfo();

                    if (office_installed_key != null && office_installed_key.Count > 0)
                    {
                        foreach (var now_key in office_installed_key)
                        {
                            string cmd_switch_cd = $"pushd \"{App.Path.Dir_SDK + @"\Activate"}\"";                  //切换至OSPP文件目录
                            string cmd_remove = $"cscript ospp.vbs /unpkey:{now_key}";
                            string result_log = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_remove})");
                            if (result_log.Contains("success"))
                            {
                                new Log($"     √ 已移除 {now_key} 激活信息。", ConsoleColor.DarkGreen);
                            }
                            else
                            {
                                new Log(result_log);    //获取失败原因
                                new Log($"     × {now_key} 对应的 Office 激活信息移除失败！", ConsoleColor.DarkRed);
                                return false;   //有一个产品卸载失败了，就直接返回 false。
                            }
                        }

                        //逐一卸载后，若都为 success，则再执行一次检查
                        office_installed_key = OfficeLocalInstall.LicenseInfo();   //再度获取list
                        if (office_installed_key == null)   //为空，视为成功
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

        /// <summary>
        /// 使用常规工具卸载安装
        /// </summary>
        internal class Uninstall
        {
            /// <summary>
            /// 强制删除 Office
            /// </summary>
            /// <returns></returns>
            internal static bool ForceDelete()
            {
                try
                {
                    new Log($"\n------> 正在执行 Office 强行清理 ...", ConsoleColor.DarkCyan);

                    //移除激活信息
                    Activate.RemoveAll();

                    //删除文件
                    try
                    {
                        //有些文件可能无法彻底删除，但如果后面复查时，不影响安装，则不会返回 false
                        var office_installed_dir = OfficeLocalInstall.GetArchiDir();
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

                    //清理注册表残余
                    try
                    {
                        //Office 有些注册表是无法清理干净的，所以用try
                        Register.DeleteItem(Registry.LocalMachine, @"SOFTWARE\Microsoft", "Office");
                    }
                    catch {/*注册表 common 项目 铁定删不掉*/}

                    //清除开始菜单
                    List<string> startmenu_list = new List<string>
                    {
                        $@"{Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)}\Programs\Microsoft Office",   //公共目录
                        $@"{Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)}\Programs\Microsoft Office",         //当前账户下目录
                    };
                    foreach (var now_dir in startmenu_list)
                    {
                        if (Directory.Exists(now_dir))
                        {
                            Directory.Delete(now_dir, true);
                        }
                    }

                    //复查是否干净了
                    var installed_info = OfficeLocalInstall.GetArchiDir();
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
            /// <summary>
            /// 使用 SaRA 完成卸载
            /// </summary>
            /// <returns></returns>
            internal static bool BySaRA()
            {
                try
                {
                    new Log($"\n------> 正在卸载 Office 冗余版本 ...", ConsoleColor.DarkCyan);
                    Thread.Sleep(1000);     //基于体验，延迟1s

                    //定义SaRA文件位置
                    string SaRA_path_root = App.Path.Dir_SDK + @"\SaRA";
                    string SaRA_path_exe = SaRA_path_root + @"\SaRACmd.exe";

                    //检查SaRA文件是否存在
                    if (!File.Exists(SaRA_path_exe))
                    {
                        new Log($"     × 目录：{SaRA_path_root} 下文件丢失，请重新下载本软件！", ConsoleColor.DarkRed);
                        return false;
                    }

                    new Log($"     >> 此过程会持续较长时间，这取决于您电脑中已安装的 Office 情况，请耐心等待 ...", ConsoleColor.DarkYellow);

                    //执行卸载命令
                    string cmd_switch_cd = $"pushd \"{SaRA_path_root}\"";             //切换至SaRA文件目录
                    string cmd_unstall = $"SaRACmd.exe -S OfficeScrubScenario -AcceptEula -Officeversion All";
                    string uninstall_result = Com_ExeOS.RunCmd($"({cmd_switch_cd})&({cmd_unstall})");
                    if (!uninstall_result.ToLower().Contains("successful"))
                    {
                        new Log(uninstall_result);
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

                    new Log($"     √ 已完成 Office 卸载操作。", ConsoleColor.DarkGreen);

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            /// <summary>
            /// 使用ODT工具卸载所有Office版本
            /// </summary>
            /// <returns></returns>
            internal static bool ByODT()
            {
                try
                {
                    new Log($"\n------> 正在卸载 Office ODT 版本 ...", ConsoleColor.DarkCyan);
                    Thread.Sleep(1000);     //基于体验，延迟1s

                    //定义ODT文件位置
                    string ODT_path_root = App.Path.Dir_SDK + @"\ODT\";
                    string ODT_path_exe = ODT_path_root + @"ODT.exe";
                    string ODT_path_xml = ODT_path_root + @"uninstall.xml";    //此文件需要新生成

                    //生成卸载xml
                    string xml_content = "<Configuration>\n  <Remove All=\"TRUE\" />\n  <Display Level=\"NONE\" AcceptEULA=\"TRUE\"/>\n</Configuration>";
                    File.WriteAllText(ODT_path_xml, xml_content);

                    //检查ODT文件是否存在
                    if (!File.Exists(ODT_path_exe) || !File.Exists(ODT_path_xml))
                    {
                        new Log($"\n     × 目录：{ODT_path_root} 下文件丢失，请重新下载本软件！", ConsoleColor.DarkRed);
                        return false;
                    }

                    new Log($"     >> 此过程大约会持续短暂的几分钟时间，请耐心等待 ...", ConsoleColor.DarkYellow);

                    //移除所有激活信息
                    if (!Activate.RemoveAll())
                    {
                        new Log("     × 移除激活信息失败！");
                        return false;
                    }

                    //执行卸载命令
                    string uninstall_args = $"/configure \"{ODT_path_xml}\"";
                    bool isUninstall = Com_ExeOS.RunExe(ODT_path_exe, uninstall_args);      //卸载

                    var reg_info = Register.GetValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIds");

                    //未正常结束卸载 OR 注册表键值不为空 时，视为卸载失败
                    if (!isUninstall || !string.IsNullOrEmpty(reg_info))
                    {
                        new Log($"     × 卸载 Office ODT 版本失败！", ConsoleColor.DarkRed);
                        return false;
                    }

                    //卸载正常 + 注册表已清空时，开始安装新版本
                    new Log($"     √ 已完成 Office 卸载操作。", ConsoleColor.DarkGreen);
                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }
        }
    }
}
