/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_FileOS.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 对文件操作的类库
    /// </summary>
    internal class Com_FileOS
    {
        /// <summary>
        /// 对xml文件操作的类库
        /// </summary>
        internal class XML
        {
            /// <summary>
            /// 通过键修改对应的值
            /// </summary>
            /// <param name="xml_path"></param>
            /// <param name="Key_str"></param>
            /// <returns></returns>
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

        /// <summary>
        /// 搜索文件的类库
        /// </summary>
        internal class ScanFiles
        {
            /// <summary>
            /// 检索最终的文件路径列表
            /// </summary>
            internal List<string> FilesList = new List<string>();

            /// <summary>
            /// 通过文件后缀扩展类型，递归查找满足条件的文件路径
            /// </summary>
            /// <param name="dirPath"></param>
            /// <param name="isRoot"></param>
            internal void GetFilesByExtension(string dirPath, string fileType = "*", bool isRoot = false)
            {
                if (Directory.Exists(dirPath))     //目录存在
                {
                    DirectoryInfo folder = new DirectoryInfo(dirPath);

                    //Console.ForegroundColor = ConsoleColor.DarkYellow;
                    //Console.Write("\r正在检索: " + folder.FullName);

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
                                //Console.ForegroundColor = ConsoleColor.DarkGray;
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
                    /*Console.ForegroundColor = ConsoleColor.DarkGray;
                    new Log("不存在: " + dirPath);*/
                }
            }
        }

        /// <summary>
        /// 转换文件不同格式，如：流、文件流等
        /// </summary>
        internal class Covert
        {
            /* - - - - - - - - - - - - - - - - - - - - - - - - 
             * Stream 和 byte[] 之间的转换
             * - - - - - - - - - - - - - - - - - - - - - - - */
            /// <summary>
            /// 将 Stream 转成 byte[]
            /// </summary>
            internal static byte[] StreamToBytes(Stream stream)
            {
                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);

                // 设置当前流的位置为流的开始
                stream.Seek(0, SeekOrigin.Begin);
                return bytes;
            }

            /// <summary>
            /// 将 byte[] 转成 Stream
            /// </summary>
            internal static Stream BytesToStream(byte[] bytes)
            {
                Stream stream = new MemoryStream(bytes);
                return stream;
            }


            /* - - - - - - - - - - - - - - - - - - - - - - - - 
             * Stream 和 文件之间的转换
             * - - - - - - - - - - - - - - - - - - - - - - - */
            /// <summary>
            /// 将 Stream 写入文件
            /// </summary>
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

            /// <summary>
            /// 从文件读取 Stream
            /// </summary>
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

        /// <summary>
        /// 文件写出类
        /// </summary>
        internal class Write
        {
            /// <summary>
            /// 将已知的所有文本一次性写入文本，并覆盖此前内容
            /// </summary>
            /// <returns></returns>
            internal static bool FromAllText(string all_text, string to_path)
            {
                try
                {
                    File.WriteAllText(to_path, all_text, Encoding.UTF8);
                    return true;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return false;
                }
            }

            /// <summary>
            /// 将Stream数据写出到本地文件
            /// </summary>
            /// <returns></returns>
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
    }
}
