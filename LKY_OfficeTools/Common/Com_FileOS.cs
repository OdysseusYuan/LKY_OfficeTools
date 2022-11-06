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
                catch /*(Exception Ex)*/
                {
                    //Console.ForegroundColor = ConsoleColor.DarkRed;
                    //new Log(Ex.Message.ToString());
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
            public void GetFilesByExtension(string dirPath, string fileType = "*", bool isRoot = false)
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

    }
}
