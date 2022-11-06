/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Aria2c.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using static LKY_OfficeTools.Lib.Lib_SelfLog;

namespace LKY_OfficeTools.SDK.Aria2c
{
    internal class Aria2c
    {
        internal static bool DownFile(string uri, string save_to)
        {
            try
            {
                //指定路径
                string aria2c_path = Environment.CurrentDirectory + @"\SDK\Aria2c\aria2c.exe";

                if (!File.Exists(aria2c_path))
                {
                    new Log("     × aria2c.exe 文件丢失！", ConsoleColor.DarkRed);
                    return false;
                }

                string file_path = new FileInfo(save_to).DirectoryName;     //保存的文件路径，不含文件名
                string filename = new FileInfo(save_to).Name;               //保存的文件名

                //设置命令行
                string aria2c_params = $"{uri} --dir=\"{file_path}\" --out=\"{filename}\" --continue=true --max-connection-per-server=5 --check-integrity=true --file-allocation=none";
                //new Log(aria2c_params);

                Com_ExeOS.RunExe(aria2c_path, aria2c_params);

                return true;
            }
            catch /*(Exception Ex)*/
            {
                //string error = Ex.Message.ToString();

                //Console.ForegroundColor = ConsoleColor.DarkRed;
                //new Log(error);
                return false;
            }
        }



    }
}
