/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_FileOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.IO;
using System.Text;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 对文件操作的类库
    /// </summary>
    internal class Com_FileOS
    {
        /// <summary>
        /// 对xml文件操作的类库
        /// </summary>
        internal class XML
        {
            /// <summary>
            /// 通过键修改对应的值
            /// </summary>
            /// <param name="xml_path"></param>
            /// <param name="Key_str"></param>
            /// <returns></returns>
            internal static bool SetValue(string xml_path, string Key_str, string new_Value)
            {
                try
                {
                    //读取文件
                    string xml_content = File.ReadAllText(xml_path, Encoding.UTF8);

                    //获得当前值
                    string current_value = Com_TextOS.GetCenterText(xml_content, $"{Key_str}=\"", "\"");

                    //替换值
                    string xml_new_content = Com_TextOS.ReplaceText(xml_content, current_value, new_Value);

                    //判断是否替换成功
                    if (string.IsNullOrEmpty(xml_new_content))
                    {
                        //Console.WriteLine("替换失败！");
                        return false;
                    }

                    //写入文件
                    File.WriteAllText(xml_path, xml_new_content, Encoding.UTF8);

                    return true;
                }
                catch (Exception Ex)
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine(Ex.Message.ToString());
                    return false;
                }
            }
        }
    }
}
