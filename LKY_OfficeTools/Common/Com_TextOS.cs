/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_TextOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

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
            catch (Exception Ex)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(Ex.Message.ToString());
                return null;
            }
        }
    }
}
