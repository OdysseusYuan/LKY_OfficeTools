/*
 *      [LKY Common Tools] Copyright (C) 2022 liukaiyuan@sjtu.edu.cn Inc.
 *      
 *      FileName : ThunderSdk.cs
 *      Developer: liukaiyuan@sjtu.edu.cn (Odysseus.Yuan)
 */

using System;
using System.Runtime.InteropServices;

namespace LKY_ThunderSdk
{
    /// <summary>
    /// API
    /// </summary>
    public static class ThunderSdk
    {
        private const string DllName = @"P:\专用工具\CommonTools\src\LKY_OfficeTools\SDK\ThunderSdk\xldl.dll";

        /// <summary>
        /// 初始化迅雷下载引擎
        /// </summary>
        /// <returns>初始化是否成功</returns>
        [DllImport(DllName, EntryPoint = "XL_Init")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool Init();

        /// <summary>
        /// 卸载迅雷下载引擎
        /// </summary>
        /// <returns>卸载是否成功</returns>
        [DllImport(DllName, EntryPoint = "XL_UnInit")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnInit();

        /// <summary>
        /// 新建下载任务
        /// </summary>
        /// <param name="param">任务参数</param>
        /// <returns>任务句柄，null表示失败</returns>
        [DllImport(DllName, EntryPoint = "XL_CreateTask", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr CreateTask([In()]DownTaskParam param);

        /// <summary>
        /// 启动下载任务
        /// </summary>
        /// <param name="task">任务句柄</param>
        /// <returns>是否成功启动任务</returns>
        [DllImport(DllName, EntryPoint = "XL_StartTask", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool StartTask(IntPtr task);

        /// <summary>
        /// 停止下载任务
        /// </summary>
        /// <param name="task">任务句柄</param>
        /// <returns>是否成功停止了任务</returns>
        [DllImport(DllName, EntryPoint = "XL_StopTask", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool StopTask(IntPtr task);

        /// <summary>
        /// 设置全局下载速度
        /// </summary>
        /// <param name="nKBps">限制速度(KB)</param>
        [DllImport(DllName, EntryPoint = "XL_SetSpeedLimit", CallingConvention = CallingConvention.Cdecl)]
        public static extern void SetSpeedLimit(int nKBps);

        /// <summary>
        /// 设置全局上传速度
        /// </summary>
        /// <param name="nTcpKBps"></param>
        /// <param name="nOtherKBps"></param>
        /// <returns></returns>
        [DllImport(DllName, EntryPoint = "XL_SetUploadSpeedLimit", CallingConvention = CallingConvention.Cdecl)]
        public static extern void SetUploadSpeedLimit(int nTcpKBps, int nOtherKBps);

        /// <summary>
        /// 查询任务信息
        /// </summary>
        /// <param name="task">任务句柄</param>
        /// <param name="taskInfo">任务信息</param>
        /// <returns>查询是否成功</returns>
        [DllImport(DllName, EntryPoint = "XL_QueryTaskInfoEx", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool QueryTaskInfoEx(IntPtr task, [Out()]TaskInfo taskInfo);

        /// <summary>
        /// 释放任务
        /// </summary>
        /// <param name="task">任务句柄</param>
        /// <returns>是否移除成功</returns>
        [DllImport(DllName, EntryPoint = "XL_DeleteTask", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool DeleteTask(IntPtr task);

        /// <summary>
        /// 删除任务数据文件
        /// </summary>
        /// <param name="param">仅需要填充SavePath和FileName两个字段</param>
        /// <returns></returns>
        [DllImport(DllName, EntryPoint = "XL_DelTempFile", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DelTempFile([In()]DownTaskParam param);

        /// <summary>
        /// 下载任务参数结构
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 1)]
        public class DownTaskParam
        {
            public int Reserved0;

            /// <summary>
            /// 下载地址
            /// </summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 2084)]
            public string TaskUrl;

            /// <summary>
            /// 引用页
            /// </summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 2084)]
            public string RefUrl;

            /// <summary>
            /// 浏览器Cookie
            /// </summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 4096)]
            public string Cookies;

            /// <summary>
            /// 本地保存文件名
            /// </summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string FileName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string Reserved1;

            /// <summary>
            /// 文件保存目录
            /// </summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string SavePath;

            public IntPtr Reserved2;

            public int Reserved3 = 0;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string Reserved4;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string Reserved5;

            /// <summary>
            /// 是否只从原始地址下载
            /// </summary>
            public bool IsOnlyOriginal = true;

            public uint Reserved6 = 5;

            /// <summary>
            /// 禁止智能命名
            /// </summary>
            public bool DisableAutoRename = false;

            /// <summary>
            /// 是否用续传
            /// </summary>
            public bool IsResume = true;

            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2048, ArraySubType = UnmanagedType.U4)]
            public uint[] Reserved7;
        }

        /// <summary>
        /// 任务状态码
        /// </summary>
        public enum TaskStatus
        {
            Noitem = 0,

            /// <summary>
            /// 下载出错
            /// </summary>
            Error,

            /// <summary>
            /// 下载已暂停
            /// </summary>
            Pause,

            /// <summary>
            /// 下载中
            /// </summary>
            Download,

            /// <summary>
            /// 下载成功
            /// </summary>
            Complete,

            /// <summary>
            /// 任务启动中
            /// </summary>
            Startpending,

            /// <summary>
            /// 任务停止中
            /// </summary>
            Stoppending
        }

        /// <summary>
        /// 错误码
        /// </summary>
        public enum ErrorCode
        {
            Unknown = 0,
            DiskCreate = 1,
            DiskWrite = 2,
            DiskRead = 3,
            DiskRename = 4,
            DiskPiecehash = 5,
            DiskFilehash = 6,
            DiskDelete = 7,
            DownInvalid = 16,
            ProxyAuthTypeUnkown = 32,
            ProxyAuthTypeFailed = 33,
            HttpmgrNotIp = 48,
            Timeout = 64,
            Cancel = 65,
            TpCrashed = 66,
            IdInvalid = 67
        }

        /// <summary>
        /// 任务信息
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public class TaskInfo
        {
            public TaskStatus State;
            public ErrorCode FailCode;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string FileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string Reserved0;
            public long TotalSize;
            public long TotalDownload;
            public float Percent;
            public int Reserved1;
            public int SrcTotal;
            public int SrcUsing;
            public int Reserved2;
            public int Reserved3;
            public int Reserved4;
            public int Reserved5;
            public long Reserved6;
            public long DonationP2P;
            public long Reserved7;
            public long DonationOrgin;
            public long DonationP2S;
            public long Reserved8;
            public long Reserved9;
            public int Speed;
            public int SpeedP2S;
            public int SpeedP2P;
            public bool IsOriginUsable;
            public float HashPercent;
            public int IsCreatingFile;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64, ArraySubType = UnmanagedType.U4)]
            public uint[] Reserved10;
        }
    }
}
