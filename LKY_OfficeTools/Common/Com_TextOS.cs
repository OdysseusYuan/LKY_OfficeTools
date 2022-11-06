/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_TextOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Security.Cryptography;
using System.Text;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 对文本操作的类
    /// </summary>
    internal class Com_TextOS
    {
        /// <summary>
        /// 获取中间文本
        /// </summary>
        /// <param name="from_text"></param>
        /// <param name="left_str"></param>
        /// <param name="right_str"></param>
        /// <returns></returns>
        internal static string GetCenterText(string from_text, string left_str, string right_str)
        {
            try
            {
                int left_str_Index = from_text.IndexOf(left_str);   //获取第一个满足左边字符的index
                from_text = from_text.Substring(left_str_Index).Remove(0, left_str.Length);    //截取左侧文本开始之后的内容（不含左侧字符串）
                int right_str_Index = from_text.IndexOf(right_str);   //获取第一个满足右边字符的index
                string result = from_text.Substring(0, right_str_Index);    //获取最终值
                return result;
            }
            catch /*(Exception Ex)*/
            {
                //Console.ForegroundColor = ConsoleColor.DarkRed;
                //new Log(Ex.Message.ToString());
                return null;
            }
        }

        ///<summary>
        ///生成随机字符串 
        ///</summary>
        ///<param name="length">目标字符串的长度</param>
        ///<param name="useNum">是否包含数字，1=包含，默认为包含</param>
        ///<param name="useLow">是否包含小写字母，1=包含，默认为包含</param>
        ///<param name="useUpp">是否包含大写字母，1=包含，默认为包含</param>
        ///<param name="useSpe">是否包含特殊字符，1=包含，默认为不包含</param>
        ///<param name="custom">要包含的自定义字符，直接输入要包含的字符列表</param>
        ///<returns>指定长度的随机字符串</returns>
        internal static string GetRandomString(int length, bool useNum, bool useLow, bool useUpp, bool useSpe, string custom)
        {
            byte[] b = new byte[4];
            new System.Security.Cryptography.RNGCryptoServiceProvider().GetBytes(b);
            Random r = new Random(BitConverter.ToInt32(b, 0));
            string s = null, str = custom;
            if (useNum == true) { str += "0123456789"; }
            if (useLow == true) { str += "abcdefghijklmnopqrstuvwxyz"; }
            if (useUpp == true) { str += "ABCDEFGHIJKLMNOPQRSTUVWXYZ"; }
            if (useSpe == true) { str += "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"; }
            for (int i = 0; i < length; i++)
            {
                s += str.Substring(r.Next(0, str.Length - 1), 1);
            }
            return s;
        }

        /// <summary>
        /// 将MD5转换为大写
        /// </summary>
        /// <param name="strPwd"></param>
        /// <returns></returns>
        internal static string Encrypt16(string strPwd)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            string t2 = BitConverter.ToString(md5.ComputeHash(UTF8Encoding.Default.GetBytes(strPwd)), 4, 8);
            t2 = t2.Replace("-", "");
            return t2.ToLower();
        }
    }
}
