/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 - 2023 OdysseusYuan@foxmail.com Inc.
 *      
 *      FileName : Lib_AppSdk.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using static LKY_OfficeTools.Common.Com_ExeOS;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.AppPath;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppMessage;
using static LKY_OfficeTools.Lib.Lib_AppState;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_AppSdk
    {
        enum SdkPackage
        {
            Activate,

            Aria2c,

            ODT,

            SaRA,
        }

        private static Dictionary<SdkPackage, Stream> SdkPackageDic
        {
            get
            {
                try
                {
                    //初始字典
                    Dictionary<SdkPackage, Stream> res_dic = new Dictionary<SdkPackage, Stream>();

                    var asm = Assembly.GetExecutingAssembly();

                    res_dic[SdkPackage.Activate] = asm.GetManifestResourceStream(AppDevelop.NameSpace_Top /* 当命名空间发生改变时，此值也需要调整 */
                                                    + $".Resource.SDK.{SdkPackage.Activate}.lotp");
                    res_dic[SdkPackage.Aria2c] = asm.GetManifestResourceStream(AppDevelop.NameSpace_Top /* 当命名空间发生改变时，此值也需要调整 */
                                                    + $".Resource.SDK.{SdkPackage.Aria2c}.lotp");
                    res_dic[SdkPackage.ODT] = asm.GetManifestResourceStream(AppDevelop.NameSpace_Top /* 当命名空间发生改变时，此值也需要调整 */
                                                    + $".Resource.SDK.{SdkPackage.ODT}.lotp");
                    res_dic[SdkPackage.SaRA] = asm.GetManifestResourceStream(AppDevelop.NameSpace_Top /* 当命名空间发生改变时，此值也需要调整 */
                                                    + $".Resource.SDK.{SdkPackage.SaRA}.lotp");
                    return res_dic;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }

        private static Dictionary<SdkPackage, string> SdkPkgPath
        {
            get
            {
                try
                {
                    //初始字典
                    Dictionary<SdkPackage, string> extra_dic = new Dictionary<SdkPackage, string>();

                    extra_dic[SdkPackage.Activate] = Documents.SDKs.SDKs_Root + $"\\{SdkPackage.Activate}.lotp";
                    extra_dic[SdkPackage.Aria2c] = Documents.SDKs.SDKs_Root + $"\\{SdkPackage.Aria2c}.lotp";
                    extra_dic[SdkPackage.ODT] = Documents.SDKs.SDKs_Root + $"\\{SdkPackage.ODT}.lotp";
                    extra_dic[SdkPackage.SaRA] = Documents.SDKs.SDKs_Root + $"\\{SdkPackage.SaRA}.lotp";

                    return extra_dic;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }

        internal static bool Initial()
        {
            try
            {
                new Log($"\n------> 正在配置 {AppAttribute.AppName} 基础组件 ...", ConsoleColor.DarkCyan);
                Thread.Sleep(1000);      //短暂间隔，提升下体验
                new Log($"     >> 此过程会持续些许时间，这取决于您的电脑硬件配置，请耐心等待 ...", ConsoleColor.DarkYellow);

                //初始化前先清理SDK目录，防止因为文件已经存在，引发解压的catch
                Clean();

                //释放文件
                if (SdkPackageDic == null)
                {
                    throw new Exception("读取 SDK 内存资源失败！");
                }
                foreach (var now_pkg in SdkPackageDic)
                {
                    string pkg_path = SdkPkgPath[now_pkg.Key];
                    bool isToDisk = Com_FileOS.Write.FromStream(now_pkg.Value, pkg_path);
                    if (!isToDisk)
                    {
                        //写出异常，抛出
                        throw new IOException($"无法写出 SDK 文件 {pkg_path} 到硬盘！");
                    }

                    //无异常，解压包
                    ZipFile.ExtractToDirectory(pkg_path, Documents.SDKs.SDKs_Root + $@"\{now_pkg.Key}");
                }

                new Log($"     √ 已完成 {AppAttribute.AppName} 组件配置。", ConsoleColor.DarkGreen);

                return true;
            }
            catch (IOException IO_Ex)
            {
                new Log(IO_Ex.ToString());

                //读写出现意外
                new Log($"     × 配置 {AppAttribute.AppName} 基础组件失败。请确保您的系统盘具备足够的可写空间！", ConsoleColor.DarkRed);

                //清理SDK缓存
                Clean();

                Current_StageType = ProcessStage.Finish_Fail;     //设置为失败模式

                //退出提示
                KeyMsg.Quit(-2);

                return false;
            }
            catch (UnauthorizedAccessException Au_Ex)
            {
                new Log(Au_Ex.ToString());

                //不具备读写权限
                new Log($"     × 配置 {AppAttribute.AppName} 基础组件失败。请确保您具备对 {Documents.SDKs.SDKs_Root} 目录的写入权限！", ConsoleColor.DarkRed);

                //清理SDK缓存
                Clean();

                Current_StageType = ProcessStage.Finish_Fail;     //设置为失败模式

                //退出提示
                KeyMsg.Quit(-2);

                return false;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());

                //其它未知问题
                new Log($"     × 配置 {AppAttribute.AppName} 基础组件失败，无法继续。请重新下载本软件或联系开发者！", ConsoleColor.DarkRed);

                //清理SDK缓存
                Clean();

                Current_StageType = ProcessStage.Finish_Fail;     //设置为失败模式

                //退出提示
                KeyMsg.Quit(-10);

                return false;
            }
            finally
            {
                //清理 SDK pkg文件
                var extra_sdk_list = SdkPkgPath.Values.ToList();
                foreach (var now_path in extra_sdk_list)
                {
                    if (File.Exists(now_path))
                    {
                        try
                        {
                            File.Delete(now_path);
                        }
                        catch (Exception Ex)
                        {
                            new Log(Ex.ToString());
                            new Log($"Exception: 清理 SDK 的 {now_path} 文件失败！");
                        }
                    }
                }
            }
        }

        internal static bool Clean()
        {
            try
            {
                //目录不存在时，自动返回为真
                if (!Directory.Exists(Documents.SDKs.SDKs_Root))
                {
                    return true;
                }

                Directory.Delete(Documents.SDKs.SDKs_Root, true);

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                new Log($"清理 SDK 目录失败！");
                return false;
            }
        }

        internal static List<string> Process_List
        {
            get
            {
                try
                {
                    Stream sdk_processes_res = Assembly.GetExecutingAssembly().
                    GetManifestResourceStream(AppDevelop.NameSpace_Top /* 当命名空间发生改变时，此值也需要调整 */
                    + ".Resource.SDK.SDK_Processes.list");
                    StreamReader sdk_processes_sr = new StreamReader(sdk_processes_res);
                    string sdk_processes = sdk_processes_sr.ReadToEnd();
                    if (!string.IsNullOrWhiteSpace(sdk_processes))
                    {
                        List<string> sdk_processes_list = new List<string>();
                        string[] p_info = sdk_processes.Replace("\r", "").Split('\n');      //分割出进程数组
                        if (p_info != null && p_info.Length > 0)
                        {
                            foreach (var now_process in p_info)
                            {
                                sdk_processes_list.Add(now_process);
                            }

                            return sdk_processes_list;
                        }
                    }
                    return null;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }

        internal static bool KillAllSdkProcess(KillExe.KillMode mode)
        {
            try
            {
                //轮询结束每个进程（不等待）
                foreach (var now_p in Process_List)
                {
                    KillExe.ByExeName(now_p, mode, false);
                }

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }
    }
}
