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
using System.Windows.Forms;

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
        internal static string Visit_WebRequest(string url)
        {
            try
            {
                Uri uri = new Uri(url);
                WebRequest myReq = WebRequest.Create(uri);
                WebResponse result = myReq.GetResponse();
                Stream receviceStream = result.GetResponseStream();
                StreamReader readerOfStream = new StreamReader(receviceStream, Encoding.UTF8);
                string strHTML = readerOfStream.ReadToEnd();
                readerOfStream.Close();
                receviceStream.Close();
                result.Close();
                return strHTML;
            }
            catch /*(Exception Ex)*/
            {
                //Console.ForegroundColor = ConsoleColor.DarkYellow;
                //Console.WriteLine($"      * 运行统计异常！");
                return null;
            }
        }
    }
}
