/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 LiuKaiyuan. All rights reserved.
 *      
 *      FileName : Com_FileOS.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    internal class Com_FileOS
    {
        internal class XML
        {
            internal static bool SetValue(string xml_path, string Key_str, string new_Value)
            {
                try
                {
                    //读取文件
                    string xml_content = File.ReadAllText(xml_path, Encoding.UTF8);

                    //获得当前值
                    string current_value = Com_TextOS.GetCenterText(xml_content, $"{Key_str}=\"", "\"");

                    //替换值
                    string xml_new_content = xml_content.Replace(current_value, new_Value);

                    //判断是否替换成功
                    if (string.IsNullOrEmpty(xml_new_content))
                    {
                        //new Log("替换失败！");
                        return false;
                    }

                    //写入文件
                    File.WriteAllText(xml_path, xml_new_content, Encoding.UTF8);

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }
        }

        internal class ScanFiles
        {
            internal List<string> FilesList = new List<string>();

            internal void GetFilesByExtension(string dirPath, string fileType = "*", bool isRoot = false)
            {
                if (Directory.Exists(dirPath))     //目录存在
                {
                    DirectoryInfo folder = new DirectoryInfo(dirPath);

                    //获取当前目录下的文件名
                    foreach (FileInfo file in folder.GetFiles())
                    {
                        //如果没有限制文件后缀名，或者满足了特定后缀名，开始写入List
                        if (fileType == "*" || file.Extension == (fileType))
                        {
                            //扫描到的文件添加到列表中
                            FilesList.Add(file.FullName);
                        }
                    }

                    //如果是根目录先排除掉 回收站目录
                    if (isRoot)
                    {
                        foreach (DirectoryInfo dir in folder.GetDirectories())
                        {
                            if (dir.FullName.Contains("$RECYCLE.BIN") || dir.FullName.Contains("System Volume Information"))
                            {
                                //Console.ForegroundColor = ConsoleColor.Gray;
                                //new Log("跳过: " + dir.FullName);
                            }
                            else
                            {
                                //new Log("----->: " + dir.FullName);
                                GetFilesByExtension(dir.FullName, fileType);
                            }
                        }
                    }
                    else
                    {
                        //遍历下一个子目录
                        foreach (DirectoryInfo subFolders in folder.GetDirectories())
                        {
                            //new Log(subFolders.FullName);
                            GetFilesByExtension(subFolders.FullName, fileType);
                        }
                    }
                }
                else
                {
                    /*Console.ForegroundColor = ConsoleColor.Gray;
                    new Log("不存在: " + dirPath);*/
                }
            }
        }

        internal class Covert
        {
            /* - - - - - - - - - - - - - - - - - - - - - - - - 
             * Stream 和 byte[] 之间的转换
             * - - - - - - - - - - - - - - - - - - - - - - - */
            internal static byte[] StreamToBytes(Stream stream)
            {
                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);

                // 设置当前流的位置为流的开始
                stream.Seek(0, SeekOrigin.Begin);
                return bytes;
            }

            internal static Stream BytesToStream(byte[] bytes)
            {
                Stream stream = new MemoryStream(bytes);
                return stream;
            }


            /* - - - - - - - - - - - - - - - - - - - - - - - - 
             * Stream 和 文件之间的转换
             * - - - - - - - - - - - - - - - - - - - - - - - */
            internal static void StreamToFile(Stream stream, string fileName)
            {
                // 把 Stream 转换成 byte[]
                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);
                // 设置当前流的位置为流的开始
                stream.Seek(0, SeekOrigin.Begin);

                // 把 byte[] 写入文件
                FileStream fs = new FileStream(fileName, FileMode.Create);
                BinaryWriter bw = new BinaryWriter(fs);
                bw.Write(bytes);
                bw.Close();
                fs.Close();
            }

            internal static Stream FileToStream(string fileName)
            {
                // 打开文件
                FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                // 读取文件的 byte[]
                byte[] bytes = new byte[fileStream.Length];
                fileStream.Read(bytes, 0, bytes.Length);
                fileStream.Close();
                // 把 byte[] 转换成 Stream
                Stream stream = new MemoryStream(bytes);
                return stream;
            }
        }

        internal class Write
        {
            internal static bool TextToFile(string file_path, string content, bool is_append = true)
            {
                try
                {
                    // 创建文件所在目录
                    Directory.CreateDirectory(new FileInfo(file_path).DirectoryName);

                    //非追加模式，并且已经存在目标文件，则先删除原有文件
                    if (!is_append && File.Exists(file_path))
                    {
                        File.Delete(file_path);
                    }

                    var fs = new FileStream(file_path, FileMode.OpenOrCreate, FileAccess.Write);
                    var sw = new StreamWriter(fs);
                    sw.BaseStream.Seek(0, SeekOrigin.End);
                    sw.WriteLine(content);

                    sw.Flush();
                    sw.Close();
                    fs.Close();

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            internal static bool FromStream(Stream stream, string to_path)
            {
                try
                {
                    byte[] Save = Covert.StreamToBytes(stream);

                    //创建文件所在目录
                    Directory.CreateDirectory(new FileInfo(to_path).DirectoryName);

                    FileStream fsObj = new FileStream(to_path, FileMode.Create);
                    fsObj.Write(Save, 0, Save.Length);
                    fsObj.Close();

                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }
        }

        internal class Info
        {
            internal static string GetHash(string file_path)
            {
                try
                {
                    //文件不存在，返回null
                    if (!File.Exists(file_path))
                    {
                        return null;
                    }

                    //为了防止文件被占用，无法校验，这里先拷贝一份副本，校验副本的值
                    string file_copied = DateTime.Now.ToFileTime().ToString() + ".hash";
                    File.Copy(file_path, file_copied, true);
                    if (!File.Exists(file_copied))
                    {
                        return null;        //副本文件拷贝失败，返回null
                    }

                    //var hash = SHA256.Create();
                    //var hash = MD5.Create();
                    var hash = SHA1.Create();
                    var stream = new FileStream(file_copied, FileMode.Open);
                    byte[] hashByte = hash.ComputeHash(stream);
                    stream.Close();

                    //校验完成，删除副本
                    if (File.Exists(file_copied))
                    {
                        File.Delete(file_copied);
                    }

                    return BitConverter.ToString(hashByte).Replace("-", "");
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return null;
                }
            }
        }
    }
}
