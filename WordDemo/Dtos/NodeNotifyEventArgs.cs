using System;
using WordDemo.Dtos;

namespace Events
{
    public class NodeNotifyEventArgs : EventArgs, IProgressModel
    {
        /// <summary>
        /// 1:进度消息  2：错误消息
        /// </summary>
        public int Type { get; set; } = 1;
        public string Title { get; set; }
        public int CurrentStep { get; set; }
        public int TotalSteps { get; set; }
        public string Message { get; set; }
    }
}
