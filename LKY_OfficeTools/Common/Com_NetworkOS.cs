/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_NetworkOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 网络处理类库
    /// </summary>
    internal class Com_NetworkOS
    {
        /// <summary>
        /// 网络环境检查
        /// </summary>
        internal class Check
        {
            //导入判断网络是否连接的 .dll
            [DllImport("wininet.dll", EntryPoint = "InternetGetConnectedState")]
            //判断网络状况的方法,返回值true为连接，false为未连接
            private extern static bool InternetGetConnectedState(out int conState, int reder);

            /// <summary>
            /// 网络联网检查
            /// </summary>
            internal static bool IsConnected
            {
                get
                {
                    return InternetGetConnectedState(out int n, 0);
                }
            }
        }

        /// <summary>
        /// IP 相关类库
        /// </summary>
        internal class IP
        {
            /// <summary>
            /// 获得自身IP地址以及网络信息
            /// </summary>
            /// <returns>返回IP和查询地址</returns>
            internal static string GetMyIP_Info()
            {
                try
                {
                    string ip = GetMyIP();
                    if (string.IsNullOrEmpty(ip))
                    {
                        //没获取到IP，则返回未知
                        ip = "ip_unknow";
                    }

                    string ip_location = IP2Location();
                    if (string.IsNullOrEmpty(ip_location))
                    {
                        //没获取到归属地，则返回未知
                        ip_location = "location_unknow";
                    }

                    return $"{ip} ({ip_location})";
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    //意外失败，返回error
                    return "ip_info_error!";
                }
            }

            /// <summary>
            /// 获得自身IP地址
            /// </summary>
            /// <returns></returns>
            internal static string GetMyIP()
            {
                try
                {
                    //截取服务器列表
                    string ip_server_info = Com_TextOS.GetCenterText(AppJson.Info, "\"IP_Check_Url_List\": \"", "\"");

                    //遍历获取ip
                    if (!string.IsNullOrEmpty(ip_server_info))
                    {
                        List<string> ip_server = new List<string>(ip_server_info.Split(';'));
                        foreach (var now_server in ip_server)
                        {
                            //获取成功时结束，否则遍历获取 
                            string my_ip_page = Com_WebOS.Visit_WebClient(now_server.Replace(" ", ""));    //替换下无用的空格字符后访问Web
                            string my_ip = GetIPFromHtml(my_ip_page);
                            if (string.IsNullOrEmpty(my_ip))
                            {
                                //获取失败则继续遍历
                                continue;
                            }
                            else
                            {
                                return my_ip;
                            }
                        }
                        //始终没获取到IP，则返回null
                        return null;
                    }
                    else
                    {
                        //获取失败时，使用默认值
                        string req_url = "http://www.net.cn/static/customercare/yourip.asp";
                        string my_ip_page = Com_WebOS.Visit_WebClient(req_url);
                        return GetIPFromHtml(my_ip_page);
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }

            /// <summary>
            /// 将 IP 映射到归属地
            /// </summary>
            /// <returns></returns>
            internal static string IP2Location()
            {
                try
                {
                    string req_url = $"https://{DateTime.UtcNow.Year}.ip138.com";

                    //访问页面
                    string ip_page = Com_WebOS.Visit_WebClient(req_url);

                    //页面非空判断
                    if (string.IsNullOrEmpty(ip_page))
                    {
                        throw new Exception();
                    }

                    //拆解html
                    string ip_location = Com_TextOS.GetCenterText(ip_page, "来自：", "</p>").Replace("\r", "").Replace("\n", "");

                    //解析非空判断
                    if (string.IsNullOrEmpty(ip_location))
                    {
                        throw new Exception();
                    }

                    //无任何问题，直接返回该值
                    return ip_location;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }

            /// <summary>
            /// 从html中通过正则找到ip信息(只支持ipv4地址)
            /// </summary>
            /// <param name="pageHtml"></param>
            /// <returns></returns>
            private static string GetIPFromHtml(string pageHtml)
            {
                try
                {
                    //验证ipv4地址
                    string reg = @"(?:(?:(25[0-5])|(2[0-4]\d)|((1\d{2})|([1-9]?\d)))\.){3}(?:(25[0-5])|(2[0-4]\d)|((1\d{2})|([1-9]?\d)))";
                    string ip = "";
                    Match m = Regex.Match(pageHtml, reg);
                    if (m.Success)
                    {
                        ip = m.Value;
                    }
                    return ip;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }


        }
    }
}
