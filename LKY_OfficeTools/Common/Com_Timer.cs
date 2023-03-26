/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2023 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : Com_Timer.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Threading;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    /// <summary>
    /// 时间器相关类库
    /// </summary>
    internal class Com_Timer
    {
        /// <summary>
        /// 倒计时器子类
        /// </summary>
        internal class Countdown_Timer
        {
            /// <summary>
            /// 剩余时间（秒）
            /// </summary>
            internal int Remaining_Time
            { get; set; }

            /// <summary>
            /// 是否运行倒计时
            /// </summary>
            internal bool isRun
            { get; set; }

            /// <summary>
            /// 开启一个计时器
            /// </summary>
            /// <param name="total_time">预计倒计时的总时间（秒）</param>
            internal void Start(int total_time)
            {
                try
                {
                    //此处不放入线程，否则使用isRun判断是否在运行时，会导致isRun状态无法判断
                    Remaining_Time = total_time;
                    isRun = true;

                    Thread time_t = new Thread(() =>
                    {
                        Update();

                        //迭代完成后，自动停止运行
                        isRun = false;
                    });

                    //time_t.SetApartmentState(ApartmentState.STA);
                    time_t.Start();
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return;
                }
            }

            /// <summary>
            /// 迭代计算
            /// </summary>
            /// <returns></returns>
            void Update()
            {
                try
                {
                    //倒计时不为0，且给出运行指令
                    if (Remaining_Time > 0 & isRun)
                    {
                        Thread.Sleep(1000);
                        Remaining_Time--;
                        Update();  //轮询继续
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return;
                }
            }

            /// <summary>
            /// 暂停计时器
            /// </summary>
            internal void Pause()
            {
                try
                {
                    isRun = false;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return;
                }
            }

            /// <summary>
            /// 恢复计时器
            /// </summary>
            internal void Continue()
            {
                try
                {
                    if (Remaining_Time != 0)
                    {
                        isRun = true;
                    }
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return;
                }
            }

            /// <summary>
            /// 停止计时器
            /// </summary>
            internal void Stop()
            {
                try
                {
                    Pause();
                    Remaining_Time = 0;
                }
                catch (Exception Ex)
                {
                    new Log(Ex.ToString());
                    return;
                }
            }
        }
    }
}
