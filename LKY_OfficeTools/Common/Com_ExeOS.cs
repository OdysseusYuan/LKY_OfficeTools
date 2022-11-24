/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_ExeOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security;
using static LKY_OfficeTools.Lib.Lib_AppLog;

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
                Console.ForegroundColor = ConsoleColor.Gray;

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
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
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

        /// <summary>
        /// 在服务模式下显式的运行 exe
        /// </summary>
        internal class Service
        {
            #region Structures
            [StructLayout(LayoutKind.Sequential)]
            internal struct SECURITY_ATTRIBUTES
            {
                internal int Length;
                internal IntPtr lpSecurityDescriptor;
                internal bool bInheritHandle;
            }
            [StructLayout(LayoutKind.Sequential)]
            internal struct STARTUPINFO
            {
                internal int cb;
                internal string lpReserved;
                internal string lpDesktop;
                internal string lpTitle;
                internal uint dwX;
                internal uint dwY;
                internal uint dwXSize;
                internal uint dwYSize;
                internal uint dwXCountChars;
                internal uint dwYCountChars;
                internal uint dwFillAttribute;
                internal uint dwFlags;
                internal short wShowWindow;
                internal short cbReserved2;
                internal IntPtr lpReserved2;
                internal IntPtr hStdInput;
                internal IntPtr hStdOutput;
                internal IntPtr hStdError;
            }
            [StructLayout(LayoutKind.Sequential)]
            internal struct PROCESS_INFORMATION
            {
                internal IntPtr hProcess;
                internal IntPtr hThread;
                internal uint dwProcessId;
                internal uint dwThreadId;
            }
            #endregion
            #region Enumerations
            enum TOKEN_TYPE : int
            {
                TokenPrimary = 1,
                TokenImpersonation = 2
            }
            enum SECURITY_IMPERSONATION_LEVEL : int
            {
                SecurityAnonymous = 0,
                SecurityIdentification = 1,
                SecurityImpersonation = 2,
                SecurityDelegation = 3,
            }
            enum WTSInfoClass
            {
                InitialProgram,
                ApplicationName,
                WorkingDirectory,
                OEMId,
                SessionId,
                UserName,
                WinStationName,
                DomainName,
                ConnectState,
                ClientBuildNumber,
                ClientName,
                ClientDirectory,
                ClientProductId,
                ClientHardwareId,
                ClientAddress,
                ClientDisplay,
                ClientProtocolType
            }
            #endregion

            #region Constants
            internal const int TOKEN_DUPLICATE = 0x0002;
            internal const uint MAXIMUM_ALLOWED = 0x2000000;
            internal const int CREATE_NEW_CONSOLE = 0x00000010;
            internal const int IDLE_PRIORITY_CLASS = 0x40;
            internal const int NORMAL_PRIORITY_CLASS = 0x20;
            internal const int HIGH_PRIORITY_CLASS = 0x80;
            internal const int REALTIME_PRIORITY_CLASS = 0x100;
            #endregion

            #region Win32 API Imports
            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool CloseHandle(IntPtr hSnapshot);
            [DllImport("kernel32.dll")]
            static extern uint WTSGetActiveConsoleSessionId();
            [DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true), SuppressUnmanagedCodeSecurityAttribute]
            static extern bool WTSQuerySessionInformation(System.IntPtr hServer, int sessionId, WTSInfoClass wtsInfoClass, out System.IntPtr ppBuffer, out uint pBytesReturned);
            [DllImport("advapi32.dll", EntryPoint = "CreateProcessAsUser", SetLastError = true, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
            internal extern static bool CreateProcessAsUser(IntPtr hToken, string lpApplicationName, string lpCommandLine, ref SECURITY_ATTRIBUTES lpProcessAttributes,
                ref SECURITY_ATTRIBUTES lpThreadAttributes, bool bInheritHandle, int dwCreationFlags, IntPtr lpEnvironment,
                String lpCurrentDirectory, ref STARTUPINFO lpStartupInfo, out PROCESS_INFORMATION lpProcessInformation);
            [DllImport("kernel32.dll")]
            static extern bool ProcessIdToSessionId(uint dwProcessId, ref uint pSessionId);
            [DllImport("advapi32.dll", EntryPoint = "DuplicateTokenEx")]
            internal extern static bool DuplicateTokenEx(IntPtr ExistingTokenHandle, uint dwDesiredAccess,
                ref SECURITY_ATTRIBUTES lpThreadAttributes, int TokenType,
                int ImpersonationLevel, ref IntPtr DuplicateTokenHandle);
            [DllImport("kernel32.dll")]
            static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);
            [DllImport("advapi32", SetLastError = true), SuppressUnmanagedCodeSecurityAttribute]
            static extern bool OpenProcessToken(IntPtr ProcessHandle, int DesiredAccess, ref IntPtr TokenHandle);
            #endregion
            
            internal static string GetCurrentActiveUser()
            {
                IntPtr hServer = IntPtr.Zero, state = IntPtr.Zero;
                uint bCount = 0;
                // obtain the currently active session id; every logged on user in the system has a unique session id  
                uint dwSessionId = WTSGetActiveConsoleSessionId();
                string domain = string.Empty, userName = string.Empty;
                if (WTSQuerySessionInformation(hServer, (int)dwSessionId, WTSInfoClass.DomainName, out state, out bCount))
                {
                    domain = Marshal.PtrToStringAuto(state);
                }
                if (WTSQuerySessionInformation(hServer, (int)dwSessionId, WTSInfoClass.UserName, out state, out bCount))
                {
                    userName = Marshal.PtrToStringAuto(state);
                }
                return string.Format("{0}\\{1}", domain, userName);
            }

            /// <summary>  
            /// Launches the given application with full admin rights, and in addition bypasses the Vista UAC prompt  
            /// </summary>  
            /// <param name="applicationName">The name of the application to launch</param>  
            /// <param name="procInfo">Process information regarding the launched application that gets returned to the caller</param>  
            /// <returns></returns>  
            internal static bool StartProcessAndByPassUAC(string applicationName, string command, out PROCESS_INFORMATION procInfo)
            {
                uint winlogonPid = 0;
                IntPtr hUserTokenDup = IntPtr.Zero, hPToken = IntPtr.Zero, hProcess = IntPtr.Zero;
                procInfo = new PROCESS_INFORMATION();
                // obtain the currently active session id; every logged on user in the system has a unique session id  
                uint dwSessionId = WTSGetActiveConsoleSessionId();
                // obtain the process id of the winlogon process that is running within the currently active session  
                Process[] processes = Process.GetProcessesByName("winlogon");
                foreach (Process p in processes)
                {
                    if ((uint)p.SessionId == dwSessionId)
                    {
                        winlogonPid = (uint)p.Id;
                    }
                }
                // obtain a handle to the winlogon process  
                hProcess = OpenProcess(MAXIMUM_ALLOWED, false, winlogonPid);
                // obtain a handle to the access token of the winlogon process  
                if (!OpenProcessToken(hProcess, TOKEN_DUPLICATE, ref hPToken))
                {
                    CloseHandle(hProcess);
                    return false;
                }
                // Security attibute structure used in DuplicateTokenEx and CreateProcessAsUser  
                // I would prefer to not have to use a security attribute variable and to just   
                // simply pass null and inherit (by default) the security attributes  
                // of the existing token. However, in C# structures are value types and therefore  
                // cannot be assigned the null value.  
                SECURITY_ATTRIBUTES sa = new SECURITY_ATTRIBUTES();
                sa.Length = Marshal.SizeOf(sa);
                // copy the access token of the winlogon process; the newly created token will be a primary token  
                if (!DuplicateTokenEx(hPToken, MAXIMUM_ALLOWED, ref sa, (int)SECURITY_IMPERSONATION_LEVEL.SecurityIdentification, (int)TOKEN_TYPE.TokenPrimary, ref hUserTokenDup))
                {
                    CloseHandle(hProcess);
                    CloseHandle(hPToken);
                    return false;
                }
                // By default CreateProcessAsUser creates a process on a non-interactive window station, meaning  
                // the window station has a desktop that is invisible and the process is incapable of receiving  
                // user input. To remedy this we set the lpDesktop parameter to indicate we want to enable user   
                // interaction with the new process.  
                STARTUPINFO si = new STARTUPINFO();
                si.cb = (int)Marshal.SizeOf(si);
                si.lpDesktop = @"winsta0\default"; // interactive window station parameter; basically this indicates that the process created can display a GUI on the desktop  
                                                   // flags that specify the priority and creation method of the process  
                int dwCreationFlags = NORMAL_PRIORITY_CLASS | CREATE_NEW_CONSOLE;
                // create a new process in the current user's logon session  
                bool result = CreateProcessAsUser(hUserTokenDup,        // client's access token  
                                                applicationName,        // file to execute  
                                                command,                // command line  
                                                ref sa,                 // pointer to process SECURITY_ATTRIBUTES  
                                                ref sa,                 // pointer to thread SECURITY_ATTRIBUTES  
                                                false,                  // handles are not inheritable  
                                                dwCreationFlags,        // creation flags  
                                                IntPtr.Zero,            // pointer to new environment block   
                                                null,                   // name of current directory   
                                                ref si,                 // pointer to STARTUPINFO structure  
                                                out procInfo            // receives information about new process  
                                                );
                // invalidate the handles  
                CloseHandle(hProcess);
                CloseHandle(hPToken);
                CloseHandle(hUserTokenDup);
                return result; // return the result  
            }
        }
    }
}
