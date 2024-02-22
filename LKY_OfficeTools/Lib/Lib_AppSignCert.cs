/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 LiuKaiyuan. All rights reserved.
 *      
 *      FileName : Lib_AppSignCert.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppInfo.AppPath;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    internal class Lib_AppSignCert
    {
        internal Lib_AppSignCert()
        {
            try
            {
                if (!AlreadyImported("12EA025393C6D19347EFB7C71313A9DD"))
                {
                    string cer_filename = "PublisherCert.cer";
                    string cer_path = Documents.Temp + $"\\{cer_filename}";

                    //cer文件不存在时，写出到运行目录
                    if (!File.Exists(cer_path))
                    {
                        Assembly assm = Assembly.GetExecutingAssembly();
                        Stream istr = assm.GetManifestResourceStream(AppDevelop.NameSpace_Top /* 当命名空间发生改变时，词值也需要调整 */ + $".Resource.{cer_filename}");
                        Com_FileOS.Write.FromStream(istr, cer_path);
                    }

                    //导入证书
                    ImportCert(cer_path);
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }

        internal static bool AlreadyImported(string serial_number)
        {
            try
            {
                X509Store store2 = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
                store2.Open(OpenFlags.MaxAllowed);
                X509Certificate2Collection certs = store2.Certificates.Find(X509FindType.FindBySerialNumber, serial_number, false);  //用序列号作为检索
                store2.Close();

                if (certs.Count == 0 || certs[0].NotAfter < DateTime.Now)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return false;
            }
        }

        internal static bool ImportCert(string cert_filepath, string cert_password = null)
        {
            try
            {
                //根据是否有密码决定导入方式
                X509Certificate2 certificate = null;
                if (string.IsNullOrEmpty(cert_password))
                {
                    //无密码
                    certificate = new X509Certificate2(cert_filepath);
                }
                else
                {
                    //有密码
                    certificate = new X509Certificate2(cert_filepath, cert_password);
                }

                certificate.FriendlyName = AppAttribute.Developer + " DigiCert";   //设置有友好名字

                X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
                store.Open(OpenFlags.ReadWrite);
                store.Remove(certificate);              //先移除
                store.Add(certificate);
                store.Close();

                //安装后删除
                File.Delete(cert_filepath);

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
