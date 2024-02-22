/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 LiuKaiyuan. All rights reserved.
 *      
 *      FileName : Com_TextOS.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.Security.Cryptography;
using System.Text;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    internal class Com_TextOS
    {
        internal static string GetCenterText(string str, string strLeft, string strRight)
        {
            try
            {
                if (str == null || str.Length == 0) return "";
                if (strLeft != "")
                {
                    int indexLeft = str.IndexOf(strLeft);//左边字符串位置
                    if (indexLeft < 0) return "";
                    indexLeft = indexLeft + strLeft.Length;//左边字符串长度
                    if (strRight != "")
                    {
                        int indexRight = str.IndexOf(strRight, indexLeft);//右边字符串位置
                        if (indexRight < 0) return "";
                        return str.Substring(indexLeft, indexRight - indexLeft);//indexRight - indexLeft是取中间字符串长度
                    }
                    else return str.Substring(indexLeft, str.Length - indexLeft);//取字符串右边
                }
                else
                {
                    //取字符串左边
                    int indexRight = str.IndexOf(strRight);
                    if (indexRight <= 0) return "";
                    else return str.Substring(0, indexRight);
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return null;
            }
        }

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

        internal static string Encrypt16(string strPwd)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            string t2 = BitConverter.ToString(md5.ComputeHash(UTF8Encoding.Default.GetBytes(strPwd)), 4, 8);
            t2 = t2.Replace("-", "");
            return t2.ToLower();
        }

        internal static int GetStringTimes(string str, string scan_str)
        {
            try
            {
                int index = 0;
                int count = 0;
                while ((index = str.IndexOf(scan_str, index)) != -1)
                {
                    count++;
                    index += scan_str.Length;
                }
                return count;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return 0;
            }
        }
    }
}
