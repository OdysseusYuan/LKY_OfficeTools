/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// App 公共信息类库
    /// </summary>
    internal class Lib_AppInfo
    {
        /// <summary>
        /// 路径类
        /// </summary>
        internal class Path
        {
            /// <summary>
            /// APP 文档根目录
            /// </summary>
            internal static string Dir_Document = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\LKY Office Tools";

            /// <summary>
            /// APP 日志存储目录
            /// </summary>
            internal static string Dir_Log = $"{Dir_Document}\\Logs";

            /// <summary>
            /// APP 临时文件夹目录
            /// </summary>
            internal static string Dir_Temp = $"{Dir_Document}\\Temp";
        }

        /// <summary>
        /// 版权类
        /// </summary>
        internal class Copyright
        {
            /// <summary>
            /// 开发者拼音全拼
            /// </summary>
            internal const string Developer = "LiuKaiyuan";
        }

        /// <summary>
        /// 用于方便开发的相关类库
        /// </summary>
        internal class Develop
        {
            /// <summary>
            /// 全局父级（顶级）命名空间
            /// </summary>
            internal const string NameSpace_Top = "LKY_OfficeTools";
        }
    }
}
