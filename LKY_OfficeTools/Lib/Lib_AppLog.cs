/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppLog.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.IO;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 日志类库
    /// </summary>
    internal class Lib_AppLog
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
            /// 运行错误的截屏文件列表
            /// </summary>
            internal static List<string> error_screen_path = new List<string>();

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
                        string datatime_format = DateTime.Now.ToString("s").Replace("T", "_").Replace(":", "-");

                        //为空时，创建日志路径
                        if (string.IsNullOrEmpty(log_filepath))
                        {
                            string file_name = datatime_format + ".log";
                            log_filepath = $"{Lib_AppInfo.Path.Dir_Log}\\{file_name}";
                        }

                        //目录不存在时创建目录
                        Directory.CreateDirectory(new FileInfo(log_filepath).DirectoryName);

                        //文件不存在时创建&写入
                        StreamWriter file = new StreamWriter(log_filepath, append: true);
                        file.WriteLine($"{datatime_format}, {str.Replace("\n", "")}");     //替换换行符
                        file.Close();

                        //出现错误时，截屏保存
                        if (str.Contains("×"))
                        {
                            string err_filename = datatime_format + ".png";
                            err_filename = $"{Lib_AppInfo.Path.Dir_Log}\\{err_filename}";
                            if (Com_SystemOS.Screen.CaptureToSave(err_filename))
                            {
                                error_screen_path.Add(err_filename);
                            }
                        }
                    }
                }
                catch
                {
                    return;
                }
            }

            /// <summary>
            /// 清理所有日志及错误文件
            /// </summary>
            /// <returns></returns>
            internal static bool Clean()
            {
                try
                {
                    //清理日志
                    if (log_filepath != null)
                    {
                        try
                        {
                            File.Delete(log_filepath);
                        }
                        catch { }
                    }

                    //清理错误截屏
                    if (error_screen_path.Count > 0)
                    {
                        foreach (var now_file in error_screen_path)
                        {
                            try
                            {
                                File.Delete(now_file);
                            }
                            catch { }
                        }
                    }

                    //清理整个Log文件夹
                    if (Directory.Exists(Lib_AppInfo.Path.Dir_Log))
                    {
                        try
                        {
                            Directory.Delete(Lib_AppInfo.Path.Dir_Log, true);
                        }
                        catch { }
                    }

                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
    }
}
