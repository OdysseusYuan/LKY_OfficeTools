/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_WebOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.IO;
using System.Net;
using System.Text;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 访问网页类库
    /// </summary>
    internal class Com_WebOS
    {
        /// <summary>
        /// 读取网页链接的内容，并非以浏览器的方式加载。
        /// 一般适用于读取网页的纯文本文件
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static string Visit_WebClient(string url, Encoding encoding = null)
        {
            try
            {
                if (encoding == null)
                {
                    encoding = Encoding.UTF8;
                }

                using (WebClient WC = new WebClient())
                {
                    WC.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.84 Safari/537.36");
                    WC.Credentials = CredentialCache.DefaultCredentials;//获取或设置用于向Internet资源的请求进行身份验证的网络凭据
                    Byte[] pageData = WC.DownloadData(url); //从指定网站下载数据
                    return encoding.GetString(pageData);
                }
            }
            catch /*(Exception Ex)*/
            {
                //Console.ForegroundColor = ConsoleColor.DarkYellow;
                //new Log(Ex);
                return null;
            }
        }
    }
}
