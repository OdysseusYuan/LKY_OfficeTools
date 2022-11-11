/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Aria2c.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.SDK.Aria2c
{
    /// <summary>
    /// Aria2c 类库
    /// </summary>
    internal class Aria2c
    {
        /// <summary>
        /// 使用 Aria2c 下载单个文件。
        /// 返回值：1~下载完成；0~因aria2c.exe丢失导致无法下载；-1~因为非已知的原因下载失败。
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="save_to"></param>
        /// <returns></returns>
        internal static int DownFile(string uri, string save_to)
        {
            try
            {
                //指定路径
                string aria2c_path = Environment.CurrentDirectory + @"\SDK\Aria2c\aria2c.exe";

                if (!File.Exists(aria2c_path))
                {
                    new Log($"     × {aria2c_path} 文件丢失！", ConsoleColor.DarkRed);
                    return 0;
                }

                string file_path = new FileInfo(save_to).DirectoryName;     //保存的文件路径，不含文件名
                string filename = new FileInfo(save_to).Name;               //保存的文件名

                //设置命令行
                string aria2c_params = $"{uri} --dir=\"{file_path}\" --out=\"{filename}\" --continue=true --max-connection-per-server=5 --check-integrity=true --file-allocation=none";
                //new Log(aria2c_params);

                Com_ExeOS.RunExe(aria2c_path, aria2c_params);

                return 1;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return -1;
            }
        }



    }
}
