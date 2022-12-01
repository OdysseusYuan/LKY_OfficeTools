/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_InstallerOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using WindowsInstaller;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 对 MSI Installer 操作的类库
    /// </summary>
    internal class Com_InstallerOS
    {
        /// <summary>
        /// MSI文件属性。
        /// 具体参考：https://learn.microsoft.com/zh-cn/windows/win32/msi/property-reference
        /// </summary>
        internal enum MsiInfoType
        {
            /// <summary>
            /// 产品名称，也就是“主题”值
            /// </summary>
            ProductName,

            /// <summary>
            /// 产品ID，一般类似 {50ae0af4-6ff4-4d45-97cc-ac734201ba3c}
            /// </summary>
            ProductCode,

            /// <summary>
            /// 产品的版本
            /// </summary>
            ProductVersion
        }

        /// <summary>
        /// 获取 MSI 文件属性信息
        /// </summary>
        /// <param name="msi_path"></param>
        /// <param name="msi_info"></param>
        /// <returns></returns>
        internal static string GetProductInfo(string msi_path, MsiInfoType msi_info)
        {
            try
            {
                Type oType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
                Installer inst = Activator.CreateInstance(oType) as Installer;
                Database DB = inst.OpenDatabase(msi_path, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);

                //生成要查询的内容
                string str = $" SELECT * FROM Property WHERE Property = '{msi_info}' ";

                View thisView = DB.OpenView(str);
                thisView.Execute();
                Record thisRecord = thisView.Fetch();
                string result = thisRecord.get_StringData(2);

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
