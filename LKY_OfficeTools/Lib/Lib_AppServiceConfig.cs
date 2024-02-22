/*
 *      [LKY Office Tools] Copyright (C) 2022 - 2024 LiuKaiyuan Inc.
 *      
 *      FileName : Lib_AppServiceConfig.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using System.ServiceProcess;
using static LKY_OfficeTools.Common.Com_ServiceOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_AppServiceConfig
    {
        internal static void Setup()
        {
            try
            {
                //无论手动、被动模式，只要安装了服务，就自动更新之
                if (Com_ServiceOS.Query.IsCreated(AppAttribute.ServiceName))
                {
                    AddOrUpdate();
                }
                else
                {
                    //未安装过。根据模式不同，给出判断

                    //手动模式，提示添加
                    if (Current_RunMode == RunMode.Manual)
                    {
                        //让用户选择是否添加自检服务
                        if (KeyMsg.Choose("为保证 Office 始终处于最新正版，即将添加 Office 自动更新/正版激活 服务。"))
                        {
                            AddOrUpdate();
                        }
                        else
                        {
                            new Log($"     × 您已拒绝添加 Office 自动更新/正版激活 服务，若要重新添加，请再次运行本软件！", ConsoleColor.DarkRed);
                            return;
                        }
                    }
                    //被动模式，自动安装服务信息
                    else if (Current_RunMode == RunMode.Passive)
                    {
                        AddOrUpdate();
                    }
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }

        enum Add_Result
        {
            Add_Success,

            Add_Fail,

            Add_Error,

            Update_Success,

            Update_Fail,

            Update_Error,

            Unknow,
        }

        static Add_Result AddOrUpdate()
        {
            try
            {
                //构建服务基本信息
                string serv_name = AppAttribute.ServiceName;
                string serv_execmd = AppPath.Documents.Services.ServiceAutorun_Exe + " /service";
                string serv_displayname = AppAttribute.ServiceDisplayName;
                string serv_desc = "用于检测、下载、安装和激活正版 Office 的重要服务。如果此服务被禁用，这台计算机的用户将无法获取 Office 更新，也无法使其处于最新正版激活状态。";

                //判断是否安装
                if (Com_ServiceOS.Query.IsCreated(serv_name))
                {
                    //已安装服务
                    new Log($"\n------> 正在更新 {AppAttribute.ServiceDisplayName} 服务 ...", ConsoleColor.DarkCyan);

                    //校验其属性是否一致

                    //修改 exe 运行信息
                    if (!Com_ServiceOS.Query.CompareBinPath(serv_name, serv_execmd))
                    {
                        Com_ServiceOS.Config.Modify.BinPath(serv_name, serv_execmd);
                    }

                    //修改 DisplayName
                    var service_info = Com_ServiceOS.Query.GetService(serv_name);      //无需判断是否为空，IsCreated() 已经判断过
                    if (service_info.DisplayName != serv_displayname)
                    {
                        //DisplayName 不同时，修改之
                        Com_ServiceOS.Config.Modify.DisplayName(serv_name, serv_displayname);
                    }

                    //修改 描述信息
                    if (!Com_ServiceOS.Query.CompareDescription(serv_name, serv_desc))
                    {
                        Com_ServiceOS.Config.Modify.Description(serv_name, serv_desc);
                    }

                    //信息修改完成后，需要更新服务对应的文件
                    string serv_filepath = AppPath.Documents.Services.ServiceAutorun_Exe;
                    if (File.Exists(serv_filepath))
                    {
                        //运行的文件和服务目录下的文件 哈希值 一致时，跳过替换旧版本，否则要升级替换旧版本（常用于更新场景下）
                        if (Com_FileOS.Info.GetHash(AppPath.Executer) == Com_FileOS.Info.GetHash(serv_filepath))
                        {
                            //哈希值一致，说明，服务文件和当前运行的文件是相同的，无需替换
                            new Log($"     √ 无需升级服务，已刷新 {AppAttribute.ServiceDisplayName} 信息。", ConsoleColor.DarkGreen);
                            return Add_Result.Update_Success;
                        }

                        //文件哈希值不一样，开始替换升级
                        /* 替换逻辑
                         * 1、如果运行路径 = 服务路径，【因为服务配置发生在更新之后，相等时，文件已经是最新的，达到了预期（更新服务文件）的目的。不做任何操作】
                         * 2、运行路径 != 服务路径，【移动旧文件到trash目录，拷贝运行路径文件到服务路径】
                         */
                        if (AppPath.Executer != serv_filepath)      //用户在非服务目录下运行
                        {
                            //move方式，解决服务启动状态中，文件无法被替换的问题。（停止服务会导致自身进程被结束）
                            string dest_filepath = AppPath.Documents.Services.ServicesTrash + $"\\{DateTime.Now.ToFileTime()}.old";
                            Directory.CreateDirectory(Path.GetDirectoryName(dest_filepath));       //创建计划移动到的目录
                            File.Move(serv_filepath, dest_filepath);

                            //复制自身文件到服务目录下
                            File.Copy(AppPath.Executer, serv_filepath, true);
                        }
                    }
                    else
                    {
                        //如果服务安装了，但文件丢失了，此处将自身文件 copy 到服务专用目录
                        Directory.CreateDirectory(Path.GetDirectoryName(serv_filepath));        //先创建服务目录，否则copy文件会异常
                        File.Copy(AppPath.Executer, serv_filepath, true);
                    }

                    new Log($"     √ 已更新 {AppAttribute.ServiceDisplayName} 服务。", ConsoleColor.DarkGreen);
                    return Add_Result.Update_Success;
                }
                else
                {
                    //未安装服务时，添加服务
                    new Log($"\n------> 正在安装 Office 自动更新/正版激活 服务 ...", ConsoleColor.DarkCyan);

                    var create_result = Com_ServiceOS.Config.Create(serv_name, serv_execmd, serv_displayname, serv_desc);

                    //判断是否成功安装
                    if (create_result)
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(AppPath.Documents.Services.ServiceAutorun_Exe));    //先创建服务目录，否则copy文件会异常
                        File.Copy(AppPath.Executer, AppPath.Documents.Services.ServiceAutorun_Exe, true);                   //将当前文件复制到服务专用文件夹

                        new Log($"     √ 已安装 Office 自动更新/正版激活 服务。", ConsoleColor.DarkGreen);
                        return Add_Result.Add_Success;
                    }
                    else
                    {
                        new Log($"     × Office 自动更新/正版激活 服务安装失败！若要添加，请重新运行本软件。", ConsoleColor.DarkRed);
                        return Add_Result.Add_Fail;
                    }
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                new Log($"     × 因特殊原因，服务配置失败！如有问题，可联系开发者。", ConsoleColor.DarkRed);
                return Add_Result.Unknow;
            }
        }

        internal static bool Start()
        {
            try
            {
                //服务没有被创建时，不运行
                if (!Com_ServiceOS.Query.IsCreated(AppAttribute.ServiceName))
                {
                    return false;
                }

                //启动服务
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                        new Lib_AppServiceHub()
                };
                ServiceBase.Run(ServicesToRun);

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }

        internal static bool Stop()
        {
            try
            {
                //服务存在时，展示停止文字提示
                if (Com_ServiceOS.Query.IsCreated(AppAttribute.ServiceName))
                {
                    new Log($"\n------> 正在停止 {AppAttribute.ServiceDisplayName} 服务 ...", ConsoleColor.DarkCyan);

                    //停止服务
                    if (Com_ServiceOS.Action.Stop(AppAttribute.ServiceName))
                    {
                        new Log($"     √ 已停止 {AppAttribute.ServiceDisplayName} 服务。", ConsoleColor.DarkGreen);
                        return true;
                    }
                    else
                    {
                        new Log($"     × 无法停止 {AppAttribute.ServiceDisplayName} 服务。", ConsoleColor.DarkGreen);
                        return false;
                    }
                }
                else
                {
                    //服务不存在，直接返回真
                    return true;
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }

        internal static void RestartSelf()
        {
            try
            {
                //自身服务名称
                string serv_name = AppAttribute.ServiceName;

                //未创建服务，不能停止！
                if (!Query.IsCreated(serv_name))
                {
                    throw new Exception($"重启服务 {serv_name} 时失败。未找到该服务！");
                }

                //已安装服务，开始重启
                string cmd = $"(net stop {serv_name})&(net start {serv_name})";
                string result = Com_ExeOS.Run.Cmd(cmd);
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }
    }
}
