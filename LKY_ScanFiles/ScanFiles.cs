/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : ScanFiles.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.IO;

namespace LKY_ScanFiles
{
    internal class ScanFiles
    {
        static void Main(string[] args)
        {
            UsingFun("D:\\");
            UsingFun("E:\\");
            UsingFun("F:\\");
            UsingFun("G:\\");
            UsingFun("H:\\");
            UsingFun("I:\\");
            UsingFun("J:\\");
            UsingFun("K:\\");
            UsingFun("L:\\");
            UsingFun("M:\\");
            UsingFun("N:\\");
            UsingFun("O:\\");
            UsingFun("P:\\");
            UsingFun("Q:\\");
            UsingFun("R:\\");
            UsingFun("S:\\");

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("finished!");

            Console.ReadKey();
        }

        /// <summary>
        /// 调用函数
        /// </summary>
        /// <param name="DestDir"></param>
        private static void UsingFun(string DestDir)
        {
            tools tool = new tools();

            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine($"正在检索 {DestDir} 目录 ...");

            tool.GetDir(DestDir, true);

            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine($"共找到 {tool.index_number} 个文件");

            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine(DestDir + " done!\n");
        }
    }

    class tools
    {
        public int index_number = 0;

        public void GetDir(string dirPath, bool isRoot = false)
        {
            if (Directory.Exists(dirPath))     //目录存在
            {
                DirectoryInfo folder = new DirectoryInfo(dirPath);

                //Console.ForegroundColor = ConsoleColor.DarkYellow;
                //Console.Write("\r正在检索: " + folder.FullName);

                //获取当前目录下的文件名
                foreach (FileInfo file in folder.GetFiles("~$*.*"))
                {
                    index_number += 1;
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine(index_number.ToString("000000") + ": " + file.FullName);
                }

                //如果是根目录先排除掉 回收站目录
                if (isRoot)
                {
                    foreach (DirectoryInfo dir in folder.GetDirectories())
                    {
                        if (dir.FullName.Contains("$RECYCLE.BIN") || dir.FullName.Contains("System Volume Information"))
                        {
                            Console.ForegroundColor = ConsoleColor.DarkGray;
                            Console.WriteLine("跳过: " + dir.FullName);
                        }
                        else
                        {
                            //Console.WriteLine("----->: " + dir.FullName);
                            GetDir(dir.FullName);
                        }
                    }
                }
                else
                {
                    //遍历下一个子目录
                    foreach (DirectoryInfo subFolders in folder.GetDirectories())
                    {
                        //Console.WriteLine(subFolders.FullName);
                        GetDir(subFolders.FullName);
                    }
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("不存在: " + dirPath);
            }
        }
    }
}
