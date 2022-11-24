/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppServiceConfig.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.ServiceProcess;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 配置服务的类库
    /// </summary>
    internal class Lib_AppServiceConfig
    {
        /// <summary>
        /// 用户手动添加一个服务
        /// </summary>
        /// <returns></returns>
        internal static void AddByUser()
        {
            try
            {
                //仅在手动模式下才提示添加
                if (Current_RunMode == RunMode.Manual)
                {
                    //让用户选择是否添加自检服务
                    if (KeyMsg.Choose("为保证 Office 始终处于最新正版，即将添加 Office 自动更新/正版激活 服务。"))
                    {
                        Add();
                    }
                    else
                    {
                        new Log($"     × 您已拒绝添加 Office 自动更新/正版激活 服务，若要重新添加，请再次运行本软件！", ConsoleColor.DarkRed);
                        return;
                    }
                }
                //被动模式下，自动安装（如果已安装，则自动跳过）
                else if (Current_RunMode == RunMode.Passive)
                {
                    Add();
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }

        /// <summary>
        /// 为 APP 添加服务
        /// </summary>
        /// <returns>1 安装成功；0 安装失败；-1 安装异常</returns>
        internal static int Add()
        {
            try
            {
                //添加服务
                var create_result = Com_ServiceOS.Config.Create(
                    AppAttribute.ServiceName,
                    AppPath.Executer + " /service",
                    AppAttribute.ServiceDisplayName,
                    "用于检测、下载、安装和激活正版 Office 的重要服务。如果此服务被禁用，这台计算机的用户将无法获取 Office 更新，也无法使其处于最新正版激活状态。");

                //判断是否成功安装
                if (create_result)
                {
                    new Log($"     √ 已成功安装 Office 自动更新/正版激活 服务。", ConsoleColor.DarkGreen);
                    return 1;
                }
                else
                {
                    new Log($"     × Office 自动更新/正版激活 服务安装失败！若要添加，请重新运行本软件。", ConsoleColor.DarkRed);
                    return 0;
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return -1;
            }
        }

        /// <summary>
        /// 让 APP 以服务方式启动
        /// </summary>
        /// <returns></returns>
        internal static bool Start()
        {
            try
            {
                //服务没有被创建时，不运行
                if (!Com_ServiceOS.Config.IsCreated(AppAttribute.ServiceName))
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

        /// <summary>
        /// 停止 APP 对应的服务
        /// </summary>
        /// <returns></returns>
        internal static bool Stop()
        {
            try
            {
                //服务存在时，展示停止文字提示
                if (Com_ServiceOS.Config.IsCreated(AppAttribute.ServiceName))
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
    }
}
