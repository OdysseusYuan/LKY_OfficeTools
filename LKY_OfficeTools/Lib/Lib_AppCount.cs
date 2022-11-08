/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppCount.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using static LKY_OfficeTools.Common.Com_SystemOS;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 用于程序统计计数的类库
    /// </summary>
    internal class Lib_AppCount
    {
        /// <summary>
        /// 发送信息类库
        /// </summary>
        internal class PostInfo
        {
            /// <summary>
            /// 回收启动信息
            /// </summary>
            /// <returns></returns>
            internal static bool Running()
            {
                try
                {
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

                    //运行模式
                    string run_mode;
#if (DEBUG)
                    run_mode = "Debug";
#else
                    run_mode = "Release";
#endif

                    //标题
#if (DEBUG)
                    string title = $"[LKY OfficeTools 启动通知 ({run_mode})]";
#else
                    string title = $"[LKY OfficeTools 启动通知]";
#endif

                    //获取软件列表
                    string soft_info = null;
                    var soft_list = SoftWare.GetList();
                    foreach (var now_soft in soft_list)
                    {
                        soft_info = $"{now_soft}<br /> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp {soft_info}";
                    }

                    string content =
                        $"<font color = black>v{Assembly.GetExecutingAssembly().GetName().Version} 于 {Environment.MachineName} ({Environment.UserName}) 启动。</font><br /><br />" +
                        $"<font color=green><b>****************************** 【详细信息】 ******************************</b></font><br /><br />" +
                         $"<font color = red>【发送时间】：</font>{DateTime.Now}<br /><br />" +
                         $"<font color = red>【反馈类型】：</font>软件例行启动<br /><br />" +
                         $"<font color = red>【软件版本】：</font>v{Assembly.GetExecutingAssembly().GetName().Version} ({run_mode})<br /><br />" +
                         $"<font color = red>【启动路径】：</font>{Process.GetCurrentProcess().MainModule.FileName}<br /><br />" +
                         $"<font color = red>【系统环境】：</font>{system_ver}<br /><br />" +
                         $"<font color = red>【机器名称】：</font>{Environment.MachineName} ({Environment.UserName})<br /><br />" +
                         $"<font color = red>【网络地址】：</font>{Com_NetworkOS.IP.GetMyIP_Info()}<br /><br />" +
                         $"<font color = red>【软件列表】：</font>{soft_info}<br /><br />" +
                         /*$"<font color = red>【关联类库】：</font>{bug_class}<br /><br />" +
                         $"<font color = red>【错误提示】：</font>{bug_msg}<br /><br />" +
                         $"<font color = red>【代码位置】：</font>{bug_code}<br /><br />" +*/
                         "<font color=green><b>****************************** 【反馈结束】 ******************************</b></font>";

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

                    //开始发送
                    bool send_result = Com_EmailOS.Send_Account("LKY Software FeedBack", PostTo,
                        title, content, null, SMTPHost, SMTPuser, SMTPpass);

                    //判断是否发送成功
                    if (send_result)
                    {
                        new Log($"     >> 初始化完成 {new Random().Next(51, 70)}% ...", ConsoleColor.DarkYellow);
                        return true;
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
                catch /*(Exception Ex)*/
                {
                    new Log($"      * 已跳过非必要流程 ...", ConsoleColor.DarkMagenta);
                    //new Log(Ex);
                    //Console.ReadKey();
                    return false;
                }
            }

            /// <summary>
            /// 回收结束时的日志信息
            /// </summary>
            /// <returns></returns>
            internal static bool Finish()
            {
                try
                {
                    new Log($"\n------> 正在进行 冗余清理，请勿关闭或重启电脑 ...", ConsoleColor.DarkCyan);

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

                    //运行模式
                    string run_mode;
#if (DEBUG)
                    run_mode = "Debug";
#else
                    run_mode = "Release";
#endif

#if (DEBUG)
                    string title = $"[LKY OfficeTools 完成通知 ({run_mode})]";
#else
                    string title = $"[LKY OfficeTools 完成通知]";
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
                        $"<font color=green><b>****************************** 【详细信息】 ******************************</b></font><br /><br />" +
                         $"<font color = red>【发送时间】：</font>{DateTime.Now}<br /><br />" +
                         $"<font color = red>【反馈类型】：</font>软件自然地完成运行<br /><br />" +
                         $"<font color = red>【软件版本】：</font>v{Assembly.GetExecutingAssembly().GetName().Version} ({run_mode})<br /><br />" +
                         $"<font color = red>【启动路径】：</font>{Process.GetCurrentProcess().MainModule.FileName}<br /><br />" +
                         $"<font color = red>【系统环境】：</font>{system_ver}<br /><br />" +
                         $"<font color = red>【机器名称】：</font>{Environment.MachineName} ({Environment.UserName})<br /><br />" +
                         $"<font color = red>【网络地址】：</font>{Com_NetworkOS.IP.GetMyIP_Info()}<br /><br />" +
                         $"<font color = red>【软件列表】：</font>{soft_info}<br /><br />" +
                         /*$"<font color = red>【关联类库】：</font>{bug_class}<br /><br />" +
                         $"<font color = red>【错误提示】：</font>{bug_msg}<br /><br />" +
                         $"<font color = red>【代码位置】：</font>{bug_code}<br /><br />" +*/
                         "<font color=green><b>****************************** 【反馈结束】 ******************************</b></font>";

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

                    //初始文件列表
                    List<string> file_list = new List<string>();

                    //判断Log是否有错误
                    MailPriority priority = MailPriority.Normal;                        //一个初始的消息
                    if (File.Exists(Log.log_filepath))
                    {
                        file_list.Add(Log.log_filepath);

                        //判断是否有错误
                        if (Log.error_screen_path.Count > 0)
                        {
                            //有错误信息，设置为高优信息
                            priority = MailPriority.High;

                            //附加错误文件
                            file_list.AddRange(Log.error_screen_path);
                        }
                    }

                    //开始回收
                    bool send_result = Com_EmailOS.Send_Account("LKY Software FeedBack", PostTo,
                        title, content, file_list, SMTPHost, SMTPuser, SMTPpass, priority);

                    //判断是否发送成功
                    if (send_result)
                    {
                        return true;
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
                catch /*(Exception Ex)*/
                {
                    new Log($"      * 已跳过非必要流程 ...", ConsoleColor.DarkMagenta);
                    //new Log(Ex);
                    //Console.ReadKey();
                    return false;
                }
                finally
                {
                    //清理日志
                    Log.Clean();

                    //回显，不写日志
                    new Log($"     √ 已完成 冗余清理，部署结束。", ConsoleColor.DarkGreen, false);
                }
            }

            /*
            /// <summary>
            /// 通过访问统计链接的方式统计
            /// </summary>
            /// <returns></returns>
            internal static bool ByUrl()
            {
                try
                {
                    //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                    string info = Lib_SelfUpdate.latest_info;
                    if (string.IsNullOrEmpty(info))
                    {
                        info = Com_WebOS.Visit_WebRequest(Lib_SelfUpdate.update_json_url);
                    }

                    //获取统计地址
                    string count_url = Com_TextOS.GetCenterText(info, "\"Count_Url\": \"", "\"");
                    //为空返回异常
                    if (string.IsNullOrEmpty(count_url))
                    {
                        return false;
                    }

                    //访问统计网址，并获取返回值
                    string count_result = Com_WebOS.Visit_WebRequest(count_url);
                    //判断是否访问成功
                    if (count_result.ToLower().Contains("success"))
                    {
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        new Log($"     >> 初始化完成 {new Random().Next(51, 70)}% ...");
                        return true;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkMagenta;
                        new Log($"     >> 已跳过非必要流程 ...");
                        return false;
                    }
                }
                catch (Exception Ex)
                {
                    //Console.ForegroundColor = ConsoleColor.DarkYellow;
                    //new Log($"      * 运行统计异常！");
                    return false;
                }
            }
            */
        }
    }
}
