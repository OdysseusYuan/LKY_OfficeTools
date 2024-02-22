/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 - 2023 OdysseusYuan@foxmail.com Inc.
 *      
 *      FileName : Lib_AppInfo.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Diagnostics;
using System.IO;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_AppInfo
    {
        internal class AppAttribute
        {
            internal const string AppName = "LKY Office Tools";

            internal const string AppName_Short = "LOT";

            internal static readonly string ServiceName = ServiceDisplayName.Replace(" ", "");

            internal const string ServiceDisplayName = AppName + " Service";

            internal const string AppFilename = "LKY_OfficeTools.exe";

            internal const string AppVersion = "1.3.0.223";

            internal const string Developer = "LiuKaiyuan";
        }

        internal class AppJson
        {
            private static string AppJsonInfo = null;

            internal static string Info
            {
                get
                {
                    try
                    {
                        //内部变量非空时，直接返回其值
                        if (!string.IsNullOrEmpty(AppJsonInfo))
                        {
                            return AppJsonInfo;
                        }

#if (!DEBUG)
                        //release模式地址
                        string json_url = $"https://gitee.com/OdysseusYuan/LKY_OfficeTools/releases/download/AppInfo/LKY_OfficeTools_AppInfo.json";
#else
                        //debug模式地址
                        string json_url = $"https://gitee.com/OdysseusYuan/LOT_OnlyTest/releases/download/AppInfo_Test/test.json";
#endif

                        //获取 Json 信息
                        string result = Com_WebOS.Visit_WebClient(json_url);

                        AppJsonInfo = result;
                        return AppJsonInfo;
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                        //new Log($"     × 获取 AppJson 信息失败！", ConsoleColor.DarkRed);
                        new Log($"获取 AppJson 信息失败！");
                        return null;
                    }
                }
            }
        }

        internal class AppDevelop
        {
            internal const string NameSpace_Top = "LKY_OfficeTools";
        }

        internal class AppPath
        {
            internal static string ExecuteDir
            {
                get
                {
                    //使用 BaseDirectory 获取路径时，结尾会多一个 \ 符号，为了替换之，先在结果处增加一个斜杠，变成 \\ 结尾，然后替换掉这个 双斜杠
                    return (AppDomain.CurrentDomain.BaseDirectory + @"\").Replace(@"\\", @"");
                }
            }

            internal static string Executer
            {
                get
                {
                    return new FileInfo(Process.GetCurrentProcess().MainModule.FileName).FullName;
                }
            }

            internal class Documents
            {
                internal static string Documents_Root
                {
                    get
                    {
                        if (Lib_AppState.Must_Use_PersonalDir)
                        {
                            return $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\{AppAttribute.AppName}";            //我的文档
                        }
                        else
                        {
                            return $"{Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)}\\{AppAttribute.AppName}";     //ProgramData
                        }
                    }
                }

                internal class Update
                {
                    internal static string Update_Root
                    {
                        get
                        {
                            return $"{Documents_Root}\\Update";
                        }
                    }

                    internal static string UpdateTrash
                    {
                        get
                        {
                            if (Path.GetPathRoot(Executer) == Path.GetPathRoot(Documents_Root))
                            {
                                //相同盘符，使用默认回收站目录
                                return $"{Update_Root}\\Trash";
                            }
                            else
                            {
                                //不同盘符，使用子目录
                                return $"{ExecuteDir}\\Trash";
                            }
                        }
                    }
                }

                internal class Services
                {
                    internal static string Services_Root
                    {
                        get
                        {
                            return $"{Documents_Root}\\Services";
                        }
                    }

                    internal static string ServicesTrash
                    {
                        get
                        {
                            return $"{Services_Root}\\Trash";
                        }
                    }

                    internal static string ServiceAutorun
                    {
                        get
                        {
                            return $"{Services_Root}\\Autorun";
                        }
                    }

                    internal static string ServiceAutorun_Exe
                    {
                        get
                        {
                            return $"{ServiceAutorun}\\{AppAttribute.AppFilename}";
                        }
                    }

                    internal static string PassiveProcessInfo
                    {
                        get
                        {
                            return $"{Services_Root}\\PassiveProcess.info";
                        }
                    }
                }

                internal static string Logs
                {
                    get
                    {
                        return $"{Documents_Root}\\Logs";
                    }
                }

                internal static string Temp
                {
                    get
                    {
                        return $"{Documents_Root}\\Temp";
                    }
                }

                internal class SDKs
                {
                    internal static string SDKs_Root
                    {
                        get
                        {
                            return $"{Documents_Root}\\SDKs";
                        }
                    }

                    internal static string Activate
                    {
                        get
                        {
                            return $"{SDKs_Root}\\Activate";
                        }
                    }

                    internal static string Activate_OSPP
                    {
                        get
                        {
                            return $"{Activate}\\OSPP.VBS";
                        }
                    }
                }
            }
        }
    }
}
