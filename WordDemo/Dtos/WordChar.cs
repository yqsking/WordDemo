
namespace WordDemo.Models
{
    /// <summary>
    /// word字符
    /// </summary>
    public class WordChar
    {

        /// <summary>
        /// 页码
        /// </summary>
        public int PageNumber { get; set; }

        /// <summary>
        /// 字符顺序
        /// </summary>
        public int CharNumber { get; set; }

        /// <summary>
        /// 字符文本
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// 最小X轴坐标
        /// </summary>
        public decimal MinX { get; set; }
        
        /// <summary>
        /// 最小Y轴坐标
        /// </summary>
        public decimal MinY { get; set; }

        /// <summary>
        /// 最大y轴坐标
        /// </summary>
        public decimal MaxY { get; set; }

        /// <summary>
        /// 偏移量
        /// </summary>
        public int Offset { get; set; }

        /// <summary>
        /// 长度
        /// </summary>
        public int Length { get; set; }
    }
}
