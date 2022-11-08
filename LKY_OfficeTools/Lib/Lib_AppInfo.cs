/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            internal static string AppDocument = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\LKY Office Tools";

            /// <summary>
            /// APP 日志存储目录
            /// </summary>
            internal static string AppLog = $"{AppDocument}\\Logs";
        }
    }
}
