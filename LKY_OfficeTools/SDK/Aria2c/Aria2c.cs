/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Aria2c.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;

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
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine("     × aria2c.exe 文件丢失！");
                    return false;
                }

                string file_path = new FileInfo(save_to).DirectoryName;     //保存的文件路径，不含文件名
                string filename = new FileInfo(save_to).Name;               //保存的文件名

                //设置命令行
                string aria2c_params = $"{uri} --dir=\"{file_path}\" --out=\"{filename}\" --continue=true --max-connection-per-server=5 --check-integrity=true --file-allocation=none";
                //Console.WriteLine(aria2c_params);

                Com_ExeOS.RunExe(aria2c_path, aria2c_params);

                return true;
            }
            catch /*(Exception Ex)*/
            {
                //string error = Ex.Message.ToString();

                //Console.ForegroundColor = ConsoleColor.DarkRed;
                //Console.WriteLine(error);
                return false;
            }
        }



    }
}
