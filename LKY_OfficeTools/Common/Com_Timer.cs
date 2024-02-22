/*
 *      [LKY Common Tools] Copyright (C) 2022 - 2024 - 2023 OdysseusYuan@foxmail.com Inc.
 *      
 *      FileName : Com_Timer.cs
 *      Developer: OdysseusYuan@foxmail.com (Odysseus.Yuan)
 */

using System;
using System.Threading;
using static LKY_OfficeTools.Lib.Lib_AppLog;

namespace LKY_OfficeTools.Common
{
    internal class Com_Timer
    {
        internal class Countdown_Timer
        {
            internal int Remaining_Time
            { get; set; }

            internal bool isRun
            { get; set; }

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
