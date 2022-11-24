/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_ServiceOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.ServiceProcess;
using static LKY_OfficeTools.Common.Com_ServiceOS.Config;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 服务操作的类库
    /// </summary>
    internal class Com_ServiceOS
    {
        /// <summary>
        /// 对服务进行相应行为的类库。
        /// 主要用于启动、暂停、停止、重启服务。
        /// </summary>
        internal class Action
        {
            /// <summary>
            /// 启动一个服务
            /// </summary>
            /// <param name="serv_name"></param>
            /// <returns></returns>
            internal static bool Start(string serv_name)
            {
                try
                {
                    //未创建服务，不启动！
                    if (!IsCreated(serv_name))
                    {
                        throw new Exception($"启动服务 {serv_name} 时失败。未找到该服务！");
                    }

                    //已安装服务，开始启动
                    using (var control = new ServiceController(serv_name))
                    {
                        //没有处于运行状态的服务，才运行。
                        if (control.Status != ServiceControllerStatus.Running)
                        {
                            control.Start();
                        }
                    }
                    return true;

                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            /// <summary>
            /// 停止一个服务
            /// </summary>
            /// <param name="serv_name"></param>
            /// <returns></returns>
            internal static bool Stop(string serv_name)
            {
                try
                {
                    //未创建服务，不能停止！
                    if (!IsCreated(serv_name))
                    {
                        throw new Exception($"停止服务 {serv_name} 时失败。未找到该服务！");
                    }

                    //已安装服务，开始停止
                    using (var control = new ServiceController(serv_name))
                    {
                        //仅停止处于运行状态的服务
                        if (control.Status == ServiceControllerStatus.Running)
                        {
                            control.Stop();
                        }
                    }
                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            /// <summary>
            /// 重启一个服务
            /// </summary>
            /// <param name="serv_name"></param>
            /// <returns></returns>
            internal static bool Restart(string serv_name)
            {
                try
                {
                    //未创建服务，不能停止！
                    if (!IsCreated(serv_name))
                    {
                        throw new Exception($"重启服务 {serv_name} 时失败。未找到该服务！");
                    }

                    //已安装服务，开始重启
                    using (var control = new ServiceController(serv_name))
                    {
                        if (control.Status == ServiceControllerStatus.Running)
                        {
                            control.Continue();
                        }
                    }
                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            /// <summary>
            /// 查询一个服务的运行状态
            /// </summary>
            /// <param name="serv_name"></param>
            /// <returns>1:已停止 2:正在启动 3:正在停止 4:已运行 5:即将继续 6:即将暂停 7:已暂停 0:未知状态</returns>
            internal static int RunState(string serv_name)
            {
                try
                {
                    using (var control = new ServiceController(serv_name))
                    {
                        return (int)control.Status;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return 0;
                }
            }
        }

        /// <summary>
        /// 配置服务的类库。
        /// 主要用于增、删、改、查服务
        /// </summary>
        internal class Config
        {
            /// <summary>
            /// 创建一个服务
            /// </summary>
            /// <param name="serv_name">服务名称。在服务详情中会展示</param>
            /// <param name="serv_runpath">运行文件。服务启动时运行哪个文件</param>
            /// <param name="serv_displayname">展示名称。在服务列表中和详情信息中，均会看到这一属性</param>
            /// <param name="serv_description">描述信息。可选。</param>
            /// <returns></returns>
            internal static bool Create(string serv_name, string serv_runpath, string serv_displayname, string serv_description = null)
            {
                try
                {
                    if (IsCreated(serv_name))
                    {
                        //已经创建服务，直接返回真
                        return true;
                    }
                    else
                    {
                        //未安装，开始安装
                        string cmd_install = $"sc create \"{serv_name}\" binPath=\"{serv_runpath}\" start=auto DisplayName=\"{serv_displayname}\"";
                        var install_result = Com_ExeOS.Run.Cmd(cmd_install);

                        //非空判断，若返回值为空，则为假
                        if (string.IsNullOrEmpty(install_result))
                        {
                            throw new Exception($"创建服务 {serv_name} 时，返回值为空！");
                        }

                        //判断是否安装成功
                        if (install_result.Contains("成功") || install_result.ToLower().Contains("success"))
                        {
                            //描述信息不为空，配置描述信息
                            if (!string.IsNullOrEmpty(serv_description))
                            {
                                Modify.Description(serv_name, serv_description);
                            }

                            //无论修改描述信息是否成功，均返回成功
                            return true;
                        }
                        else
                        {
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
            /// 查询一个服务是否被创建（通过服务名称查询）
            /// </summary>
            /// <returns></returns>
            internal static bool IsCreated(string serv_name)
            {
                try
                {
                    ServiceController[] services = ServiceController.GetServices();
                    foreach (ServiceController s in services)
                    {
                        if (s.ServiceName.ToLower() == serv_name.ToLower())
                        {
                            return true;
                        }
                    }
                    return false;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            /// <summary>
            /// 删除一个服务
            /// </summary>
            /// <returns></returns>
            internal static bool Delete(string serv_name)
            {
                if (IsCreated(serv_name))
                {
                    //已经创建服务，才能删除

                    //先停止服务
                    if (!Action.Stop(serv_name))
                    {
                        throw new Exception($"服务 {serv_name} 因无法停止，导致卸载失败！");
                    }

                    string cmd_del = $"sc delete \"{serv_name}\"";
                    var del_result = Com_ExeOS.Run.Cmd(cmd_del);

                    //非空判断，若返回值为空，则为假
                    if (string.IsNullOrEmpty(del_result))
                    {
                        throw new Exception($"删除服务 {serv_name} 时，返回值为空！");
                    }

                    //判断是否删除成功
                    if (del_result.Contains("成功") || del_result.ToLower().Contains("success"))
                    {
                        return true;
                    }
                    else
                    {
                        throw new Exception($"删除服务 {serv_name} 失败！");
                    }
                }
                else
                {
                    throw new Exception($"没有找到名称为 {serv_name} 的服务，无法删除该服务！");
                }
            }

            /// <summary>
            /// 修改服务的相关信息 类库
            /// </summary>
            internal class Modify
            {
                /// <summary>
                /// 修改一个服务的描述信息
                /// </summary>
                /// <param name="serv_name">服务名称。在服务详情中会展示</param>
                /// <param name="serv_description">描述信息，将在服务列表页、详情页展示</param>
                /// <returns></returns>
                internal static bool Description(string serv_name, string serv_description)
                {
                    try
                    {
                        if (IsCreated(serv_name))
                        {
                            //已经创建服务，才进行修改
                            string cmd_modify_desc = $"sc description \"{serv_name}\"  \"{serv_description}\"";
                            var modify_result = Com_ExeOS.Run.Cmd(cmd_modify_desc);

                            //非空判断，若返回值为空，则为假
                            if (string.IsNullOrEmpty(modify_result))
                            {
                                throw new Exception($"执行修改 {serv_name} 描述信息时，返回值为空！");
                            }

                            //判断是否修改描述成功
                            if (modify_result.Contains("成功") || modify_result.ToLower().Contains("success"))
                            {
                                return true;
                            }
                            else
                            {
                                throw new Exception($"修改服务 {serv_name} 的描述信息为 {serv_description} 失败！");
                            }
                        }
                        else
                        {
                            throw new Exception($"尝试修改服务 {serv_name} 的描述信息为 {serv_description}，但无法检测到该服务已被安装！");
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
    }
}
