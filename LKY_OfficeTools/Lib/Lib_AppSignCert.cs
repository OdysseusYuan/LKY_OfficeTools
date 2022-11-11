/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Lib_AppSignCert.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using LKY_OfficeTools.Common;
using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using static LKY_OfficeTools.Lib.Lib_AppInfo;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Lib
{
    /// <summary>
    /// 数字签名的类库
    /// </summary>
    internal class Lib_AppSignCert
    {
        /// <summary>
        /// 构造函数检查证书，不存在则导入证书
        /// </summary>
        internal Lib_AppSignCert()
        {
            try
            {
                if (!check_have_cert())
                {
                    string cert_path = Lib_AppInfo.Path.Dir_Temp + "\\lky_cert.pfx";
                    string cert_key = "jae6dFktJnzURhPu6HngVhtJFNYkGVYxgLBC#rwqZJQ#drEskdP#9QPJecJf$C6uRC5w&6e9TRJPFaEFBWrRhmYDSdbMV2VwTg&";

                    //pfx文件不存在时，写出到运行目录
                    if (!File.Exists(cert_path))
                    {
                        Assembly assm = Assembly.GetExecutingAssembly();
                        Stream istr = assm.GetManifestResourceStream(Develop.NameSpace_Top /* 当命名空间发生改变时，词值也需要调整 */ + ".Resource.LKY_Cert.pfx");
                        Com_FileOS.Write.FromStream(istr, cert_path);
                    }

                    //导入证书
                    import_cert(cert_path, cert_key);
                }
            }
            catch (Exception Ex)
            {
                new Log(Ex.ToString());
                return;
            }
        }

        /// <summary>
        /// 检查本机置信区域是否已经导入了证书
        /// </summary>
        /// <returns></returns>
        internal static bool check_have_cert(string CertIssuerName = Copyright.Developer)
        {
            try
            {
                X509Store store2 = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
                store2.Open(OpenFlags.MaxAllowed);
                X509Certificate2Collection certs = store2.Certificates.Find(X509FindType.FindByIssuerName, CertIssuerName, false);  //用颁发者名字作为检索
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

        /// <summary>
        /// 导入一个证书
        /// </summary>
        /// <param name="cert_filepath"></param>
        /// <param name="cert_password"></param>
        internal static bool import_cert(string cert_filepath, string cert_password)
        {
            try
            {
                X509Certificate2 certificate = new X509Certificate2(cert_filepath, cert_password);
                certificate.FriendlyName = Copyright.Developer + " Cert";   //设置有友好名字

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
