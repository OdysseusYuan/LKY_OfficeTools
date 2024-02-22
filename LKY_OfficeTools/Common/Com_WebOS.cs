/*
 *      [LKY Office Tools] Copyright (C) 2022 - 2024 LiuKaiyuan Inc.
 *      
 *      FileName : Com_WebOS.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.Net;
using System.Text;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    internal class Com_WebOS
    {
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
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return null;
            }
        }
    }
}
