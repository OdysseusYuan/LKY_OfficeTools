/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_SelfLog.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.IO;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 日志类库
    /// </summary>
    internal class Lib_SelfLog
    {
        /// <summary>
        /// 产生记录类库
        /// </summary>
        internal class Log
        {
            /// <summary>
            /// 日志文件保存的位置
            /// </summary>
            internal static string log_filepath = null;

            /// <summary>
            /// 重载实现日志
            /// </summary>
            internal Log(string str, ConsoleColor str_color, bool is_out_file = true)
            {
                try
                {
                    Console.ForegroundColor = str_color;
                    Console.WriteLine(str);

                    //需要输出日志文件时，进行判断
                    if (is_out_file)
                    {
                        //为空时，创建日志路径
                        if (string.IsNullOrEmpty(log_filepath))
                        {
                            string file_name = DateTime.Now.ToString("s").Replace("T", "_").Replace(":", "-") + ".log";
                            log_filepath = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\LKY Office Tools\\Logs\\{file_name}";
                        }

                        //目录不存在时创建目录
                        Directory.CreateDirectory(new FileInfo(log_filepath).DirectoryName);

                        //文件不存在时创建&写入
                        StreamWriter file = new StreamWriter(log_filepath, append: true);
                        file.WriteLine($"{DateTime.Now.ToString("s").Replace("T", "_")}, {str.Replace("\n", "")}");     //替换换行符
                        file.Close();
                    }
                }
                catch
                {
                    return;
                }
            }
        }
    }
}
