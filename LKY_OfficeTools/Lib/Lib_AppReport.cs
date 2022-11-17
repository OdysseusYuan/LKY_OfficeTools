/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppReport.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Threading;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.App.State;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 用于程序 Debug 的类库
    /// </summary>
    internal class Lib_AppReport
    {
        //初始文件列表
        private static List<string> file_list = new List<string>();

        /// <summary>
        /// 点位日志上报
        /// </summary>
        /// <returns></returns>
        internal static bool Pointing(RunType point_type, bool show_info = false)
        {
            try
            {
                if (show_info && point_type != RunType.Starting)
                {
                    new Log($"\n------> 正在清理 冗余数据，请勿关闭或重启电脑 ...", ConsoleColor.DarkCyan);
                }

                //获取系统版本
                ///先获取位数
                int sys_bit = Environment.Is64BitOperatingSystem ? 64 : 32;
                string system_ver = null;
                if (OSVersion.GetPublishType() == OSVersion.OSType.Win10)
                {
                    system_ver = $"{OSVersion.OSType.Win10} ({OSVersion.GetBuildNumber(false)}) x{sys_bit} v{OSVersion.GetBuildNumber()}";
                }
                else if (OSVersion.GetPublishType() == OSVersion.OSType.Win11)
                {
                    system_ver = $"{OSVersion.OSType.Win11} ({OSVersion.GetBuildNumber(false)}) x{sys_bit} v{OSVersion.GetBuildNumber()}";
                }
                else
                {
                    system_ver = OSVersion.GetPublishType().ToString() + $" x{sys_bit}";
                }

                //点位标题
                string title = null;
                switch (point_type)
                {
                    case RunType.Starting:
                        {
                            title = $"[LKY OfficeTools 启动]";
                            break;
                        }
                    case RunType.Finish_Success:
                        {
                            title = $"[LKY OfficeTools 完成]";    //当且仅当激活成功时，视为成功
                            break;
                        }
                    case RunType.Finish_Fail:
                        {
                            title = $"[LKY OfficeTools 完成-有误]";
                            break;
                        }
                    case RunType.Interrupt:
                        {
                            title = $"[LKY OfficeTools 结束-中断]";
                            break;
                        }
                    default:
                        {
                            title = $"[LKY OfficeTools {point_type}]";
                            break;
                        }
                }

                string run_mode;
#if DEBUG
                run_mode = "Debug";
                title += $"({run_mode})";
#else
                run_mode = "Release";
#endif

                //获取软件列表
                string soft_info = null;
                var soft_list = SoftWare.GetList();
                foreach (var now_soft in soft_list)
                {
                    soft_info = $"{now_soft}<br /> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp {soft_info}";
                }

                string content =
                    $"<font color = black>v{Assembly.GetExecutingAssembly().GetName().Version} 在 {Environment.MachineName} ({Environment.UserName}) 运行结束。<br /><br />" +
                    $"<font color=green><b>------------------------------ 【简述】 ------------------------------</b></font><br /><br />" +
                     $"<font color = purple><b>【发送时间】</b>：</font>{DateTime.Now}<br /><br />" +
                     $"<font color = purple><b>【反馈类型】</b>：</font>{point_type}<br /><br />" +
                     $"<font color = purple><b>【软件版本】</b>：</font>v{Assembly.GetExecutingAssembly().GetName().Version} ({run_mode})<br /><br />" +
                     $"<font color = purple><b>【启动路径】</b>：</font>{Process.GetCurrentProcess().MainModule.FileName}<br /><br />" +
                     $"<font color = purple><b>【系统环境】</b>：</font>{system_ver}<br /><br />" +
                     $"<font color = purple><b>【机器名称】</b>：</font>{Environment.MachineName} ({Environment.UserName})<br /><br />" +
                     $"<font color = purple><b>【网络地址】</b>：</font>{Com_NetworkOS.IP.GetMyIP_Info()}<br /><br />" +
                     $"<font color = purple><b>【软件列表】</b>：</font>{soft_info}<br />";

                //非启动打点，增加log
                if (point_type != RunType.Starting)
                {
                    content += $"<font color=green><b>------------------------------ 【日志】 ------------------------------</b></font><br /><br />" +
                     $"<font color = black>{Log.log_info}</font><br />" + "";
                }

                //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                string info = Lib_AppUpdate.latest_info;
                if (string.IsNullOrEmpty(info))
                {
                    info = Com_WebOS.Visit_WebClient(Lib_AppUpdate.update_json_url);
                }

                string PostTo = Com_TextOS.GetCenterText(info, "\"Count_Feedback_To\": \"", "\"");
                //smtp
                string SMTPHost = "smtp.qq.com";
                string SMTPuser = Com_TextOS.GetCenterText(info, "\"Count_Feedback_From\": \"", "\"");
                string SMTPpass = Com_TextOS.GetCenterText(info, "\"Count_Feedback_Pwd\": \"", "\"");

                //为空抛出异常
                if (string.IsNullOrEmpty(PostTo) || string.IsNullOrEmpty(SMTPuser) || string.IsNullOrEmpty(SMTPpass))
                {
                    throw new Exception();
                }

                //判断Log是否有错误
                MailPriority priority = MailPriority.Normal;                        //一个初始的消息
                                                                                    //有错误记录，信息设置高优
                if (Log.log_info.Contains("×"))
                {
                    //有错误信息，设置为高优信息
                    priority = MailPriority.High;
                }

                //判断是否有注册表错误
                if (!string.IsNullOrEmpty(Log.reg_install_error) && File.Exists(Log.reg_install_error))
                {
                    //附加错误文件
                    file_list.Add(Log.reg_install_error);
                }

                //启动打点，增加额外内容
                if (point_type == RunType.Starting)
                {
                    if (show_info)
                    {
                        new Log($"     >> 初始化完成 {new Random().Next(31, 50)}% ...", ConsoleColor.DarkYellow);
                    }

                    //启动桌面打点上报
                    string desk_path = $"{App.Path.Dir_Log}\\running_info.jpg";
                    if (Screen.CaptureToSave(desk_path))
                    {
                        file_list.Add(desk_path);
                    }

                    if (show_info)
                    {
                        new Log($"     >> 初始化完成 {new Random().Next(51, 70)}% ...", ConsoleColor.DarkYellow);
                    }
                }

                /*
                if (File.Exists(Log.log_filepath))
                {
                    file_list.Add(Log.log_filepath);

                    //判断是否有错误
                    if (Log.error_screen_path != null)
                    {
                        //有错误信息，设置为高优信息
                        priority = MailPriority.High;

                        //附加错误文件
                        file_list.AddRange(Log.error_screen_path);
                    }
                }
                */

                //开始回收
                bool send_result = Com_EmailOS.Send_Account("LKY Software FeedBack", PostTo,
                    title, content, file_list, SMTPHost, SMTPuser, SMTPpass, priority);

                //判断是否发送成功
                if (send_result)
                {
                    //启动模式增加话术
                    if (show_info && point_type == RunType.Starting)
                    {
                        new Log($"     >> 初始化完成 {new Random().Next(71, 90)}% ...", ConsoleColor.DarkYellow);
                        Thread.Sleep(500);     //基于体验延迟一下
                    }

                    return true;
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                if (show_info)
                {
                    new Log($"      * 已跳过非必要流程 ...", ConsoleColor.DarkMagenta);
                }
                return false;
            }
            finally
            {
                //清理冗余文件
                if (file_list != null && file_list.Count > 0)
                {
                    foreach (var now_file in file_list)
                    {
                        if (File.Exists(now_file))
                        {
                            File.Delete(now_file);
                        }
                    }
                }

                if (show_info && point_type != RunType.Starting)
                {
                    //回显，不写日志
                    new Log($"     √ 已完成 冗余数据清理，部署结束。", ConsoleColor.DarkGreen, Log.Output_Type.Display);
                }
            }
        }
    }
}
