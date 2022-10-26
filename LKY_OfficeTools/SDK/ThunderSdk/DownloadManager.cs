/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : DownloadManager.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace LKY_ThunderSdk
{
    public class DownloadManager
    {
        private int _downloadSpeedLimit;
        private int _uploadSpeedLimit;

        static DownloadManager()
        {
            ThunderSdk.Init();
        }

        public List<DownFileInfo> DownloadNow { get; set; }
        private Queue<DownFileInfo> WaitingDownload { get; set; } = new Queue<DownFileInfo>();
        public List<DownFileInfo> AllDownLoad { get; set; } = new List<DownFileInfo>();
        public string SavePath { get; }

        private int DownloadSpeedLimit
        {
            get => _downloadSpeedLimit;
            set
            {
                if (value <= 0) value = int.MaxValue;
                _downloadSpeedLimit = value;
                ThunderSdk.SetSpeedLimit(value);
            }
        }

        private int UploadSpeedLimit
        {
            get => _uploadSpeedLimit;
            set
            {
                if (value <= 0) value = int.MaxValue;
                _uploadSpeedLimit = value;
                ThunderSdk.SetUploadSpeedLimit(value, value);
            }
        }
        public event EventHandler TaskInfoChanged;
        public event EventHandler TaskCompleted;
        public event EventHandler TaskNoitem;
        public event EventHandler TaskError;
        public event EventHandler TaskDownload;
        public event EventHandler TaskPause;

        public int ParallelDownloadCount { get; set; }

        public int CreateNewTask(string sourceUrl, string fileName, string savePath = "", string cookies = "", string refUrl = "", bool isOnlyOriginal = true, bool disableAutoRename = false, bool isResume = true)
        {
            if (string.IsNullOrEmpty(savePath))
                savePath = SavePath;
            var task = new DownFileInfo(sourceUrl, fileName, savePath, cookies, false, refUrl, isOnlyOriginal, disableAutoRename,
                isResume);
            var index = -1;
            lock (AllDownLoad)
            {
                AllDownLoad.Add(task);
                index = AllDownLoad.Count - 1;
            }
            lock (DownloadNow)
            {
                lock (WaitingDownload)
                {
                    if (DownloadNow.Count < ParallelDownloadCount && WaitingDownload.Count == 0)
                    {
                        DownloadNow.Add(task);
                        task.StartTask();
                    }
                    else
                    {
                        WaitingDownload.Enqueue(task);
                    }
                }
            }
            return index;
        }

        public bool PauseAllTask() => DownloadNow.All(task => ThunderSdk.StopTask(task.Id));
        public bool StartAllTask() => DownloadNow.All(task => ThunderSdk.StartTask(task.Id));


        public Thread Processor { get; set; }

        public DownloadManager(int parallelCount, string savePath = "")
        {
            ParallelDownloadCount = parallelCount;
            if (string.IsNullOrEmpty(savePath)) savePath = AppDomain.CurrentDomain.BaseDirectory;
            SavePath = savePath;
            DownloadNow = new List<DownFileInfo>();
            Processor = new Thread(() =>
            {
                while (true)
                {
                    if (DownloadNow.Count > 0)
                        if ((DownloadNow.Count == ParallelDownloadCount && WaitingDownload.Count > 0) || (DownloadNow.Count > 0 && WaitingDownload.Count == 0))
                            for (var i = 0; i < DownloadNow.Count; i++)
                            {
                                var info = DownloadNow[i];
                                if (ThunderSdk.QueryTaskInfoEx(info.Id, info.TaskInfo))
                                {
                                    OnTaskInfoChanged(info, new EventArgs());
                                    if (info.TaskInfo.State == ThunderSdk.TaskStatus.Complete)
                                    {
                                        info._wasteTime = DateTime.Now - info.CreateTime;
                                        info.TaskInfo.Percent = 1;
                                        OnTaskCompleted(info, new EventArgs());
                                        DownloadNow.Remove(info);
                                        if (WaitingDownload.Count <= 0) continue;
                                        lock (DownloadNow)
                                        {
                                            lock (WaitingDownload)
                                            {
                                                var task = WaitingDownload.Dequeue();
                                                DownloadNow.Add(task);
                                                i--;
                                                task.StartTask();
                                            }
                                            break;
                                        }
                                    }
                                    else if (info.TaskInfo.State == ThunderSdk.TaskStatus.Noitem)
                                    {
                                        info._wasteTime = DateTime.Now - info.CreateTime;
                                        OnTaskNoitem(info, new EventArgs());
                                        DownloadNow.Remove(info);
                                        if (WaitingDownload.Count <= 0) continue;
                                        lock (DownloadNow)
                                        {
                                            lock (WaitingDownload)
                                            {
                                                var task = WaitingDownload.Dequeue();
                                                DownloadNow.Add(task);
                                                i--;
                                                task.StartTask();
                                            }
                                            break;
                                        }
                                    }
                                    else if (info.TaskInfo.State == ThunderSdk.TaskStatus.Error)
                                    {
                                        info._wasteTime = DateTime.Now - info.CreateTime;
                                        OnTaskError(info, new EventArgs());
                                        DownloadNow.Remove(info);
                                        if (WaitingDownload.Count <= 0) continue;
                                        lock (DownloadNow)
                                        {
                                            lock (WaitingDownload)
                                            {
                                                var task = WaitingDownload.Dequeue();
                                                DownloadNow.Add(task);
                                                i--;
                                                task.StartTask();
                                            }
                                            break;
                                        }
                                    }
                                    else if (info.TaskInfo.State == ThunderSdk.TaskStatus.Download)
                                    {
                                        OnTaskDownload(info, new EventArgs());
                                    }
                                    else if (info.TaskInfo.State == ThunderSdk.TaskStatus.Pause)
                                    {
                                        OnTaskPause(info, new EventArgs());
                                    }
                                }
                            }
                        else if (WaitingDownload.Count > 0)
                        {
                            lock (WaitingDownload)
                            {
                                var task = WaitingDownload.Dequeue();
                                DownloadNow.Add(task);
                                task.StartTask();
                            }
                        }
                    Thread.Sleep(33);
                }
            })
            { IsBackground = true };
            Processor.Start();
        }
        private void OnTaskInfoChanged(DownFileInfo info, EventArgs eventArgs)
        {
            TaskInfoChanged?.Invoke(info, eventArgs);
        }

        private void OnTaskCompleted(DownFileInfo info, EventArgs eventArgs)
        {
            TaskCompleted?.Invoke(info, eventArgs);
        }
        public void OnTaskNoitem(DownFileInfo info, EventArgs eventArgs)
        {
            TaskNoitem?.Invoke(info, eventArgs);
        }
        public void OnTaskError(DownFileInfo info, EventArgs eventArgs)
        {
            TaskError?.Invoke(info, eventArgs);
        }
        public void OnTaskDownload(DownFileInfo info, EventArgs eventArgs)
        {
            TaskDownload?.Invoke(info, eventArgs);
        }
        public void OnTaskPause(DownFileInfo info, EventArgs eventArgs)
        {
            TaskPause?.Invoke(info, eventArgs);
        }
    }
}