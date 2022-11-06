/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_ExeOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 外部 exe 文件调用类库
    /// </summary>
    internal class Com_ExeOS
    {
        /// <summary>
        /// 启动一个外部exe
        /// </summary>
        /// <param name="runFilePath"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        internal static bool RunExe(string runFilePath, string args)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;

                Process p = new Process();
                p.StartInfo.FileName = runFilePath;             //需要启动的程序名       
                p.StartInfo.Arguments = args;                   //启动参数

                //是否使用操作系统shell启动
                p.StartInfo.UseShellExecute = false;

                //启动
                p.Start();

                //接收返回值
                //p.StandardInput.AutoFlush = true;

                //获取输出信息
                //string strOuput = p.StandardOutput.ReadToEnd();

                //等待程序执行完退出进程
                p.WaitForExit();

                p.Close();

                return true;
            }
            catch /*(Exception Ex)*/
            {
                //Console.ForegroundColor = ConsoleColor.DarkRed;
                //new Log(Ex.Message.ToString());
                return false;
            }
        }

        /// <summary>
        /// 运行CMD命令，并返回执行结果
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        internal static string RunCmd(string args)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;

                Process p = new Process();
                //设置要启动的应用程序
                p.StartInfo.FileName = "cmd.exe";

                //是否使用操作系统shell启动
                p.StartInfo.UseShellExecute = false;

                // 接受来自调用程序的输入信息
                p.StartInfo.RedirectStandardInput = true;

                //输出信息
                p.StartInfo.RedirectStandardOutput = true;

                // 输出错误
                p.StartInfo.RedirectStandardError = true;

                //不显示程序窗口
                p.StartInfo.CreateNoWindow = true;

                //启动程序
                p.Start();

                //向cmd窗口发送输入信息
                p.StandardInput.WriteLine(args + "&exit");

                p.StandardInput.AutoFlush = true;

                //获取输出信息
                string strOuput = p.StandardOutput.ReadToEnd();

                //等待程序执行完退出进程
                p.WaitForExit();
                p.Close();

                return strOuput;
            }
            catch /*(Exception Ex)*/
            {
                //Console.ForegroundColor = ConsoleColor.DarkRed;
                //new Log(Ex.Message.ToString());
                return null;
            }
        }
    }
}
