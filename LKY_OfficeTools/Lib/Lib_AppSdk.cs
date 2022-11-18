/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppSdk.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;
using static LKY_OfficeTools.Lib.Lib_AppReport;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// App SDK 类库
    /// </summary>
    internal class Lib_AppSdk
    {
        /// <summary>
        /// 资源在内存中的位置
        /// </summary>
        private static Stream sdk_package_res = Assembly.GetExecutingAssembly().
            GetManifestResourceStream(App.Develop.NameSpace_Top /* 当命名空间发生改变时，词值也需要调整 */
            + ".Resource.SDKs.pkg");

        /// <summary>
        /// 确定路径
        /// </summary>
        private static string sdk_disk_path = App.Path.SDK.Root + "\\LOT_SDKs.pkg";

        /// <summary>
        /// 初始化释放 SDK 包
        /// </summary>
        /// <returns></returns>
        internal static bool initial()
        {
            try
            {
                //初始化前先清理SDK目录，防止因为文件已经存在，引发解压的catch
                Clean();

                //释放文件
                bool isToDisk = Com_FileOS.Write.FromStream(sdk_package_res, sdk_disk_path);
                if (isToDisk)
                {
                    //解压包
                    ZipFile.ExtractToDirectory(sdk_disk_path, App.Path.SDK.Root);
                }

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                new Log($"\n\n     × 软件 SDK 文件丢失，无法继续，请重新下载本软件或联系开发者！", ConsoleColor.DarkRed);

                //清理SDK缓存
                Clean();

                Pointing(App.State.RunType.Finish_Fail);  //回收

                Environment.Exit(-1);
                return false;
            }
            finally
            {
                if (File.Exists(sdk_disk_path))
                {
                    try
                    {
                        File.Delete(sdk_disk_path);
                    }
                    catch (Exception Ex)
                    {
                        new Log(Ex.ToString());
                        new Log($"     × 清理SDK的pkg文件失败！");
                    }
                }
            }
        }

        /// <summary>
        /// 清理 SDK 包
        /// </summary>
        /// <returns></returns>
        internal static bool Clean()
        {
            try
            {
                //目录不存在时，自动返回为真
                if (!Directory.Exists(App.Path.SDK.Root))
                { 
                    return true;
                }

                Directory.Delete(App.Path.SDK.Root, true);

                return true;
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                new Log($"     × 清理SDK目录失败！");
                return false;
            }
        }
    }
}
