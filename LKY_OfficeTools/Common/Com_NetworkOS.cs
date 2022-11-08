/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_NetworkOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Lib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using static LKY_OfficeTools.Common.Com_NetworkOS.IP.ToLocation;

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
                        return "unknow";
                    }

                    //解析IP
                    ToLocation location = new ToLocation();
                    IPLocation loc = location.Get(ip);

                    string ip_info = string.Empty;
                    //归属地非空判断
                    if (!string.IsNullOrEmpty(loc.country))
                    {
                        ip_info = loc.country;
                    }

                    //运营商非空判断
                    if (!string.IsNullOrEmpty(loc.area))
                    {
                        ip_info += " - " + loc.area;
                    }

                    return $"{ip} ({ip_info})";
                }
                catch
                {
                    //意外失败，返回error
                    return "error!";
                }
            }

            /// <summary>
            /// 获得自身IP地址
            /// </summary>
            /// <returns></returns>
            internal static string GetMyIP()
            {
                //先从Update里面获取信息，如果已经访问过json，则直接用，否则重新访问
                string info = Lib_AppUpdate.latest_info;
                if (string.IsNullOrEmpty(info))
                {
                    info = Com_WebOS.Visit_WebClient(Lib_AppUpdate.update_json_url);
                }

                //截取服务器列表
                string ip_server_info = Com_TextOS.GetCenterText(info, "\"IP_Check_Url_List\": \"", "\"");

                string req_url = null;

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
                    req_url = "http://www.net.cn/static/customercare/yourip.asp";
                    string my_ip_page = Com_WebOS.Visit_WebClient(req_url);
                    return GetIPFromHtml(my_ip_page);
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
                catch
                {
                    return null;
                }
            }

            ///<summary>
            /// 提供从IP数据库搜索IP信息；
            ///</summary>
            internal class ToLocation
            {
                Stream ipFile;
                long ip;

                ///<summary>
                /// 地理位置,包括国家和地区
                ///</summary>
                internal struct IPLocation
                {
                    internal string ip, country, area;
                }

                ///<summary>
                /// 【从嵌入的dat资源读取】获取指定IP所在位置
                ///</summary>
                ///<param name="strIP">要查询的IP地址</param>
                ///<returns></returns>
                internal IPLocation Get(string strIP)
                {
                    try
                    {
                        //从资源中读取dat
                        Assembly assm = Assembly.GetExecutingAssembly();
                        ipFile = assm.GetManifestResourceStream("LKY_OfficeTools" + ".Resource." + "qqwry.dat");    //命名空间修改时，LKY_OfficeTools 也要进行相应修改。
                        IPLocation loc = new IPLocation();
                        ip = IPToLong(strIP);

                        //ip获取的long不能为异常值（-1）
                        if (ip != -1)
                        {
                            long[] ipArray = BlockToArray(ReadIPBlock());
                            long offset = SearchIP(ipArray, 0, ipArray.Length - 1) * 7 + 4;
                            ipFile.Position += offset;//跳过起始IP
                            ipFile.Position = ReadLongX(3) + 4;//跳过结束IP


                            loc.ip = strIP;

                            int flag = ipFile.ReadByte();//读取标志
                            if (flag == 1)//表示国家和地区被转向
                            {
                                ipFile.Position = ReadLongX(3);
                                flag = ipFile.ReadByte();//再读标志
                            }
                            long countryOffset = ipFile.Position;
                            loc.country = ReadString(flag).Replace("CZ88.NET", "-");     //替换掉无值时的版权信息

                            if (flag == 2)
                            {
                                ipFile.Position = countryOffset + 3;
                            }
                            flag = ipFile.ReadByte();
                            loc.area = ReadString(flag).Replace("CZ88.NET", "-");     //替换掉无值时的版权信息;

                            ipFile.Close();
                            ipFile = null;
                        }
                        return loc;
                    }
                    catch
                    {
                        return new IPLocation();
                    }
                }

                ///<summary>
                /// 将字符串形式的IP转换位long
                ///</summary>
                ///<param name="strIP"></param>
                ///<returns>异常时返回 -1</returns>
                long IPToLong(string strIP)
                {
                    try
                    {
                        //ip不为空，且有四个“.”分割的数字
                        if (!string.IsNullOrEmpty(strIP) && strIP.Split('.').Length == 4)
                        {
                            byte[] ip_bytes = new byte[8];
                            string[] strArr = strIP.Split(new char[] { '.' });
                            for (int i = 0; i < 4; i++)
                            {
                                ip_bytes[i] = byte.Parse(strArr[3 - i]);
                            }
                            return BitConverter.ToInt64(ip_bytes, 0);
                        }
                        else
                        {
                            return -1;
                        }
                    }
                    catch
                    {
                        return -1;
                    }
                }

                ///<summary>
                /// 将索引区字节块中的起始IP转换成Long数组
                ///</summary>
                ///<param name="ipBlock"></param>
                long[] BlockToArray(byte[] ipBlock)
                {
                    try
                    {
                        long[] ipArray = new long[ipBlock.Length / 7];
                        int ipIndex = 0;
                        byte[] temp = new byte[8];
                        for (int i = 0; i < ipBlock.Length; i += 7)
                        {
                            Array.Copy(ipBlock, i, temp, 0, 4);
                            ipArray[ipIndex] = BitConverter.ToInt64(temp, 0);
                            ipIndex++;
                        }
                        return ipArray;
                    }
                    catch
                    {
                        return null;
                    }
                }

                ///<summary>
                /// 从IP数组中搜索指定IP并返回其索引
                ///</summary>
                ///<param name="ipArray">IP数组</param>
                ///<param name="start">指定搜索的起始位置</param>
                ///<param name="end">指定搜索的结束位置</param>
                ///<returns></returns>
                int SearchIP(long[] ipArray, int start, int end)
                {
                    try
                    {
                        int middle = (start + end) / 2;
                        if (middle == start)
                            return middle;
                        else if (ip < ipArray[middle])
                            return SearchIP(ipArray, start, middle);
                        else
                            return SearchIP(ipArray, middle, end);
                    }
                    catch
                    {
                        return -1;
                    }
                }

                ///<summary>
                /// 读取IP文件中索引区块
                ///</summary>
                ///<returns></returns>
                byte[] ReadIPBlock()
                {
                    try
                    {
                        long startPosition = ReadLongX(4);
                        long endPosition = ReadLongX(4);
                        long count = (endPosition - startPosition) / 7 + 1;//总记录数
                        ipFile.Position = startPosition;
                        byte[] ipBlock = new byte[count * 7];
                        ipFile.Read(ipBlock, 0, ipBlock.Length);
                        ipFile.Position = startPosition;
                        return ipBlock;
                    }
                    catch
                    {
                        return null;
                    }
                }

                ///<summary>
                /// 从IP文件中读取指定字节并转换位long
                ///</summary>
                ///<param name="bytesCount">需要转换的字节数，主意不要超过8字节</param>
                ///<returns></returns>
                long ReadLongX(int bytesCount)
                {
                    try
                    {
                        byte[] _bytes = new byte[8];
                        ipFile.Read(_bytes, 0, bytesCount);
                        return BitConverter.ToInt64(_bytes, 0);
                    }
                    catch
                    {
                        return -1;
                    }
                }

                ///<summary>
                /// 从IP文件中读取字符串
                ///</summary>
                ///<param name="flag">转向标志</param>
                ///<returns></returns>
                string ReadString(int flag)
                {
                    try
                    {
                        if (flag == 1 || flag == 2)//转向标志
                            ipFile.Position = ReadLongX(3);
                        else
                            ipFile.Position -= 1;

                        List<byte> list = new List<byte>();
                        byte b = (byte)ipFile.ReadByte();
                        while (b > 0)
                        {
                            list.Add(b);
                            b = (byte)ipFile.ReadByte();
                        }
                        return Encoding.Default.GetString(list.ToArray());
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
        }
    }
}
