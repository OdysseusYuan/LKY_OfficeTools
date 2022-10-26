/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : DownFileInfo.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Threading;

namespace LKY_ThunderSdk
{
    public class DownFileInfo
    {
        private ThunderSdk.TaskInfo _taskInfo;
        public event EventHandler TaskInfoChanged;

        static DownFileInfo()
        {
            ThunderSdk.Init();
        }


        public DownFileInfo(string sourceUrl, string fileName, string savePath, string cookies = "", bool autoStart = true, string refUrl = "", bool isOnlyOriginal = true, bool disableAutoRename = false, bool isResume = true)
        {
            SourceUrl = sourceUrl;
            //FileName = fileName;
            SavePath = savePath;
            DisableAutoRename = disableAutoRename;
            IsResume = isResume;
            IsOnlyOriginal = isOnlyOriginal;
            CreateTime = DateTime.Now;
            Id = ThunderSdk.CreateTask(new ThunderSdk.DownTaskParam
            {
                TaskUrl = sourceUrl,
                FileName = fileName,
                SavePath = savePath,
                Cookies = cookies,
                RefUrl = refUrl,
                IsOnlyOriginal = isOnlyOriginal,
                DisableAutoRename = disableAutoRename,
                IsResume = isResume
            });
            if (autoStart)
                StartTaskInfoRefreasher();
        }

        //private bool _needToRefreshTaskInfo = false;

        public string SourceUrl { get; }

        public string FileName
        {
            get => TaskInfo.FileName;
        }

        public string SavePath { get; }

        public DateTime CreateTime { get; }

        internal TimeSpan _wasteTime = TimeSpan.MinValue;
        public TimeSpan WasteTime => _wasteTime == TimeSpan.MinValue ? DateTime.Now - CreateTime : _wasteTime;


        public bool IsOnlyOriginal { get; } = true;

        public bool DisableAutoRename { get; } = false;

        public bool IsResume { get; } = true;

        public ThunderSdk.TaskInfo TaskInfo
        {
            get
            {
                if (_taskInfo == null)
                    _taskInfo = new ThunderSdk.TaskInfo();
                ThunderSdk.QueryTaskInfoEx(Id, _taskInfo);
                return _taskInfo;
            }
        }

        public IntPtr Id { get; set; }

        private Thread InfoRefresher { get; set; }

        public bool StartTask() => ThunderSdk.StartTask(Id);

        public bool StopTask() => ThunderSdk.StopTask(Id);

        public bool DeleteTask() => ThunderSdk.DeleteTask(Id);

        public void StartTaskInfoRefreasher()
        {
            InfoRefresher = new Thread(() =>
                {

                    while (ThunderSdk.QueryTaskInfoEx(Id, TaskInfo))
                    {
                        OnPropertyChanged(new EventArgs());
                        if (TaskInfo.State != ThunderSdk.TaskStatus.Complete)
                        {
                            Thread.Sleep(33);
                        }
                        else
                        {
                            _wasteTime = DateTime.Now - CreateTime;
                            return;
                        }
                    }
                })
            { IsBackground = true };

            InfoRefresher.Start();
        }

        private void OnPropertyChanged(EventArgs eventArgs)
        {
            TaskInfoChanged?.Invoke(this, eventArgs);
        }
    }
}