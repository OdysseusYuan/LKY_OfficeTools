/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInstall.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
                    if (StartInstall())
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
        /// 开始安装 Office
        /// </summary>
        internal static bool StartInstall()
        {
            //定义ODT文件位置
            string ODT_path_root = Environment.CurrentDirectory + @"\SDK\ODT\";
            string ODT_path_exe = ODT_path_root + @"ODT.exe";
            string ODT_path_xml = ODT_path_root + @"config.xml";

            //检查ODT文件是否存在
            if (!File.Exists(ODT_path_exe) || !File.Exists(ODT_path_exe))
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
