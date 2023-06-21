/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppReport.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Threading;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.AppPath;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppState;

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
        internal static bool Pointing(ProcessStage point_type, bool show_info = false)
        {
            //Passive不打点
            if (Current_RunMode == RunMode.Passive)
            {
                return true;
            }

            try
            {
                if (show_info && point_type != ProcessStage.Starting)
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
                string title = $"{AppAttribute.AppName_Short}: {point_type} ({Current_RunMode})";

                //设置发布模式
                string release_mode;
#if DEBUG
                release_mode = "Debug";
                title += $" [{release_mode}]";
#else
                release_mode = "Release";
#endif

                //获取软件列表
                string soft_info = null;
                var soft_list = SoftWare.InstalledList();
                foreach (var now_soft in soft_list)
                {
                    soft_info = $"{now_soft}<br /> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp {soft_info}";
                }

                string content =
                    $"<font color = black>v{AppAttribute.AppVersion} BY {Environment.MachineName} ({Environment.UserName}) {point_type}。<br /><br />" +
                    $"<font color=green><b>------------------------------ 【简述】 ------------------------------</b></font><br /><br />" +
                     $"<font color = purple><b>【发送时间】</b>：</font>{DateTime.Now}<br /><br />" +
                     $"<font color = purple><b>【阶段类型】</b>：</font>{point_type}<br /><br />" +
                     $"<font color = purple><b>【软件版本】</b>：</font>v{AppAttribute.AppVersion} ({release_mode})<br /><br />" +
                     $"<font color = purple><b>【运行模式】</b>：</font>{Current_RunMode}<br /><br />" +
                     $"<font color = purple><b>【启动路径】</b>：</font>{Executer}<br /><br />" +
                     $"<font color = purple><b>【系统环境】</b>：</font>{system_ver}<br /><br />" +
                     $"<font color = purple><b>【机器名称】</b>：</font>{Environment.MachineName} ({Environment.UserName})<br /><br />" +
                     $"<font color = purple><b>【网络地址】</b>：</font>{Com_NetworkOS.IP.GetMyIP_Info()}<br /><br />" +
                     $"<font color = purple><b>【软件列表】</b>：</font>{soft_info}<br />";

                //非启动打点，增加log
                if (point_type != ProcessStage.Starting)
                {
                    content += $"<font color=green><b>------------------------------ 【日志】 ------------------------------</b></font><br /><br />" +
                     $"<font color = black>{Log.log_info}</font><br />" + "";
                }

                string PostTo = Com_TextOS.GetCenterText(AppJson.Info, "\"Count_Feedback_To\": \"", "\"");
                //smtp
                string SMTPHost = "smtp.qq.com";
                string SMTPuser = Com_TextOS.GetCenterText(AppJson.Info, "\"Count_Feedback_From\": \"", "\"");
                string SMTPpass = Com_TextOS.GetCenterText(AppJson.Info, "\"Count_Feedback_Pwd\": \"", "\"");

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
                if (point_type == ProcessStage.Starting)
                {
                    if (show_info)
                    {
                        new Log($"     >> 初始化完成 {new Random().Next(31, 50)}% ...", ConsoleColor.DarkYellow);
                    }

                    //启动桌面打点上报
                    string desk_path = $"{Documents.Logs}\\running_info.jpg";
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
                bool send_result = Com_EmailOS.Send_Account($"{AppAttribute.AppName} Report", PostTo,
                    title, content, file_list, SMTPHost, SMTPuser, SMTPpass, priority);

                //判断是否发送成功
                if (send_result)
                {
                    //启动模式增加话术
                    if (show_info && point_type == ProcessStage.Starting)
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
                Log.Clean();

                if (show_info && point_type != ProcessStage.Starting)
                {
                    //回显，不写日志
                    new Log($"     √ 已完成 冗余数据清理，部署结束。", ConsoleColor.DarkGreen, Log.Output_Type.Display);
                }
            }
        }
    }
}
