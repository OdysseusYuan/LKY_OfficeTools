/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_SelfCount.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Reflection;
using static LKY_OfficeTools.Common.Com_SystemOS;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 用于程序统计计数的类库
    /// </summary>
    internal class Lib_SelfCount
    {
        /// <summary>
        /// 重载实现统计
        /// </summary>
        internal Lib_SelfCount()
        {
            PostInfo.ByEmail();
        }

        /// <summary>
        /// 发送信息类库
        /// </summary>
        internal class PostInfo
        {
            /// <summary>
            /// 通过 E-mail 方式统计
            /// </summary>
            /// <returns></returns>
            internal static bool ByEmail()
            {
                try
                {
                    //获取系统版本
                    string system_ver = null;
                    if (OS.WindowsVersion() == OS.OSType.Win10)
                    {
                        system_ver = $"{OS.OSType.Win10} ({OS.Win10Version(false)}) v{OS.Win10Version()}";
                    }
                    else if (OS.WindowsVersion() == OS.OSType.Win11)
                    {
                        system_ver = $"{OS.OSType.Win11} ({OS.Win11Version(false)}) v{OS.Win10Version()}";
                    }
                    else
                    {
                        system_ver = OS.WindowsVersion().ToString();
                    }

                    //访问统计网址，并获取返回值
                    string title = $"[LKY OfficeTools 启动信息]";

                    string content =
                        $"<font color=green><b>*************** 【运行信息】 ***************</b></font><br /><br />" +
                         $"<font color = red>【发送时间】：</font>{DateTime.Now}<br /><br />" +
                         $"<font color = red>【反馈类型】：</font>软件例行启动<br /><br />" +
                         $"<font color = red>【软件版本】：</font>v{Assembly.GetExecutingAssembly().GetName().Version}<br /><br />" +
                         $"<font color = red>【系统环境】：</font>{system_ver}<br /><br />" +
                         /*$"<font color = red>【关联类库】：</font>{bug_class}<br /><br />" +
                         $"<font color = red>【错误提示】：</font>{bug_msg}<br /><br />" +
                         $"<font color = red>【代码位置】：</font>{bug_code}<br /><br />" +*/
                         "<font color=green><b>*************** 【反馈结束】 ***************</b></font>";

                    //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                    string info = Lib_SelfUpdate.latest_info;
                    if (string.IsNullOrEmpty(info))
                    {
                        info = Com_WebOS.Visit_WebRequest(Lib_SelfUpdate.update_json_url);
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
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine($"     >> 初始化完成 {new Random().Next(51, 70)}% ...");
                        return true;
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
                catch /*(Exception Ex)*/
                {
                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    Console.WriteLine($"     >> 已跳过非必要流程 ...");
                    return false;
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
                        Console.WriteLine($"     >> 初始化完成 {new Random().Next(51, 70)}% ...");
                        return true;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkMagenta;
                        Console.WriteLine($"     >> 已跳过非必要流程 ...");
                        return false;
                    }
                }
                catch (Exception Ex)
                {
                    //Console.ForegroundColor = ConsoleColor.DarkYellow;
                    //Console.WriteLine($"      * 运行统计异常！");
                    return false;
                }
            }
            */
        }
    }
}
