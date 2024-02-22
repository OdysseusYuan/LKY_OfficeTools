/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 - 2023 OdysseusYuan@foxmail.com Inc.
 *      
 *      FileName : Com_InstallerOS.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using WindowsInstaller;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    internal class Com_InstallerOS
    {
        internal enum MsiInfoType
        {
            ProductName,

            ProductCode,

            ProductVersion
        }

        internal static string GetProductInfo(string msi_path, MsiInfoType msi_info)
        {
            try
            {
                Type oType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
                if (oType == null)
                {
                    return null;
                }

                Installer inst = Activator.CreateInstance(oType) as Installer;
                if (inst == null)
                {
                    return null;
                }

                Database DB = inst.OpenDatabase(msi_path, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);
                if (DB == null)
                {
                    return null;
                }

                //生成要查询的内容
                string str = $" SELECT * FROM Property WHERE Property = '{msi_info}' ";

                View thisView = DB.OpenView(str);
                if (thisView == null)
                {
                    return null;
                }

                thisView.Execute();

                Record thisRecord = thisView.Fetch();
                if (thisRecord == null)
                {
                    return null;
                }

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
