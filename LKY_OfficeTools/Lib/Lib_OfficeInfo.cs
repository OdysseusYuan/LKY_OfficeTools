/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_OfficeInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Threading;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_OfficeInfo
    {
        /// <summary>
        /// 最新版 Office 版本信息
        /// </summary>
        internal static Version latest_version = null;

        private const string office_info_url = "https://config.office.com/api/filelist/channels";

        public static string office_file_root_url;

        /// <summary>
        /// 获取最新版本的 Office 函数
        /// </summary>
        /// <returns></returns>
        internal static Version Get_OfficeVersion()
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("\n\n------> 正在获取最新 Microsoft Office 版本 ...");

                //获取频道信息
                WebClient MyWebClient = new WebClient();
                MyWebClient.Credentials = CredentialCache.DefaultCredentials;       //获取或设置用于向Internet资源的请求进行身份验证的网络凭据
                Byte[] pageData = MyWebClient.DownloadData(office_info_url);        //从指定网站下载数据
                //string pageHtml = Encoding.Default.GetString(pageData);             //如果获取网站页面采用的是GB2312，则使用这句            
                string office_info = Encoding.UTF8.GetString(pageData);              //如果获取网站页面采用的是UTF-8，则使用这句

                if (!string.IsNullOrEmpty(office_info))
                {
                    //获取版本信息
                    string latest_info = Com_TextOS.GetCenterText(office_info, "\"PerpetualVL2021\",", "name");     //获取 2021 LTSC
                    latest_version = new Version(Com_TextOS.GetCenterText(latest_info, "latestUpdateVersion\":\"", "\"},"));              //获取版本号

                    //赋值对应的下载地址
                    office_file_root_url = Com_TextOS.GetCenterText(latest_info, "baseUrl\":\"", "\"");              //获取url

                    //Console.WriteLine(office_file_root_url);

                    return latest_version;
                }
                else
                { return null; }
            }

            catch (WebException webEx)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(webEx.Message.ToString());
                return null;
            }

        }

        /// <summary>
        /// 获取文件下载列表
        /// </summary>
        /// <returns></returns>
        internal static List<string> Get_OfficeFileList()
        {
            try
            {
                Version version_info = Get_OfficeVersion();      //获取版本

                if (version_info == null || string.IsNullOrEmpty(office_file_root_url))     //下载根地址为空时，视为失败
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine("     × 最新版本获取失败，请稍后重试！");
                    return null;
                }

                string ver = version_info.ToString();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"     √ 获取完成，最新版本：{ver}。");

                //延迟，让用户看清版本号
                Thread.Sleep(500);

                //获取文件列表
                List<string> file_list = new List<string>();
                office_file_root_url += "/office/data";
                file_list.Add($"{office_file_root_url}/v64.cab");
                file_list.Add($"{office_file_root_url}/v64_{ver}.cab");
                file_list.Add($"{office_file_root_url}/{ver}/i640.cab");
                file_list.Add($"{office_file_root_url}/{ver}/i642052.cab");
                file_list.Add($"{office_file_root_url}/{ver}/s640.cab");
                file_list.Add($"{office_file_root_url}/{ver}/s642052.cab");
                file_list.Add($"{office_file_root_url}/{ver}/stream.x64.x-none.dat");
                file_list.Add($"{office_file_root_url}/{ver}/stream.x64.zh-cn.dat");

                /*
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine($"\n------> 预计下载文件：");
                foreach (var a in file_list)
                {
                    Console.WriteLine($"      > {a}");
                }
                */

                return file_list;
            }

            catch (WebException webEx)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(webEx.Message.ToString());
                return null;
            }

        }
    }
}
