/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_ProcessOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Diagnostics;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 进程管理类库
    /// </summary>
    internal class Com_ProcessOS
    {
        /// <summary>
        /// 判断进程是否在运行
        /// </summary>
        /// <param name="exe_name">不要扩展名，例如：abc.exe，此处应填写abc</param>
        /// <returns></returns>
        internal static bool ProcessIsRun(string exe_name)
        {
            Process[] p = Process.GetProcesses();
            foreach (Process now_p in p)
            {
                if (now_p.ProcessName.Equals(exe_name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 结束指定exe名称的进程
        /// </summary>
        /// <param name="exe_name">不要扩展名，例如：abc.exe，此处应填写abc</param>
        /// <returns></returns>
        internal static bool KillProcess(string exe_name)
        {
            //先判断是否存在进程
            if (ProcessIsRun(exe_name))
            {
                try
                {
                    Process[] p = Process.GetProcessesByName(exe_name);
                    foreach (Process now_p in p)
                    {
                        now_p.Kill();
                    }
                    return true;
                }
                catch /*(Exception Ex)*/
                {
                    //Console.ForegroundColor = ConsoleColor.DarkRed;
                    //new Log(Ex.Message.ToString());
                    return false;
                }
            }
            else
            {
                //不存在时，直接返回 true
                return true;
            }
        }
    }
}
