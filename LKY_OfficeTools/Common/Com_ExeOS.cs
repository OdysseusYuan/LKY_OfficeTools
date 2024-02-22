/*
 *      [LKY Office Tools] Copyright (C) 2022 - 2024 LiuKaiyuan Inc.
 *      
 *      FileName : Com_ExeOS.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    internal class Com_ExeOS
    {
        internal class Run
        {
            internal static int Exe(string file_path, string args)
            {
                try
                {
                    Process p = new Process();
                    var result = Process(file_path, args, out p, true);     //默认等待完成

                    //是否执行了
                    if (!result)
                    {
                        throw new Exception($"执行 {file_path} 异常！");
                    }

                    return p.ExitCode;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return -920921;
                }
            }

            internal static bool Process(string file_path, string args, out Process ProcessInfo, bool WaitForExit)
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.Gray;

                    ProcessInfo = new Process();
                    ProcessInfo.StartInfo.FileName = file_path;             //需要启动的程序名       
                    ProcessInfo.StartInfo.Arguments = args;                   //启动参数

                    //是否使用操作系统shell启动
                    ProcessInfo.StartInfo.UseShellExecute = false;

                    //启动
                    ProcessInfo.Start();

                    //接收返回值
                    //p.StandardInput.AutoFlush = true;

                    //获取输出信息
                    //string strOuput = p.StandardOutput.ReadToEnd();

                    //等待程序执行完退出进程
                    if (WaitForExit)
                    {
                        ProcessInfo.WaitForExit();
                    }

                    return true;
                }
                catch (Exception Ex)
                {
                    ProcessInfo = null;
                    new Log(Ex.ToString());
                    return false;
                }
            }

            internal static string Cmd(string args)
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.Gray;

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
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }

        internal class KillExe
        {
            internal enum KillMode
            {
                Try_Friendly = 1,

                Only_Friendly = 2,

                Only_Force = 4,
            }

            internal static bool ByExeName(string exe_name, KillMode kill_mode, bool isWait)
            {
                try
                {
                    //先判断是否存在进程
                    if (Info.IsRun(exe_name))
                    {
                        Process[] p = Process.GetProcessesByName(exe_name);
                        foreach (Process now_p in p)
                        {
                            ByProcessID(now_p.Id, kill_mode, isWait);
                        }
                        return true;
                    }
                    else
                    {
                        //不存在时，直接返回 true
                        return true;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            internal static bool ByProcessID(int exe_id, KillMode kill_mode, bool isWait)
            {
                try
                {
                    //先判断是否存在进程
                    if (Info.IsRun(exe_id))
                    {
                        //进程存在，获取对象
                        Process now_p = Process.GetProcessById(exe_id);

                        switch (kill_mode)
                        {
                            case KillMode.Try_Friendly:
                                {
                                    //先尝试友好关闭
                                    if (!now_p.CloseMainWindow())
                                    {
                                        //有好关闭失败时，启用强制结束
                                        now_p.Kill();
                                    }
                                    break;
                                }
                            case KillMode.Only_Friendly:
                                {
                                    //只友好关闭
                                    now_p.CloseMainWindow();
                                    break;
                                }
                            case KillMode.Only_Force:
                                {
                                    //只强制结束
                                    now_p.Kill();
                                    break;
                                }
                        }

                        //判断是否等待进程结束
                        if (isWait)
                        {
                            //等待进程被结束
                            now_p.WaitForExit();
                        }

                        return true;    //如果不是等待结束进程，中途未出现catch时，也返回true
                    }
                    else
                    {
                        //不存在时，直接返回 true
                        return true;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }
        }

        internal class Info
        {
            internal static bool IsRun(string exe_name)
            {
                try
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
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            internal static bool IsRun(int exe_id)
            {
                try
                {
                    Process[] p = Process.GetProcesses();
                    foreach (Process now_p in p)
                    {
                        if (now_p.Id == exe_id)
                        {
                            return true;
                        }
                    }
                    return false;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            internal static List<Process> GetProcessByTitle(string window_title, bool need_equal = false)
            {
                try
                {
                    Process[] process_list = Process.GetProcesses();

                    List<Process> result = new List<Process>();
                    foreach (var now_p in process_list)
                    {
                        if (need_equal)
                        {
                            //严格相等
                            if (now_p.MainWindowTitle == window_title)
                            {
                                result.Add(now_p);
                            }
                        }
                        else
                        {
                            //包含即可
                            if (now_p.MainWindowTitle.Contains(window_title))
                            {
                                result.Add(now_p);
                            }
                        }
                    }
                    return result;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }
    }
}
