using System.Text.RegularExpressions;

namespace WordDemo.Models
{
    /// <summary>
    /// word行
    /// </summary>
    public class WordLine
    {
        /// <summary>
        /// 页码数
        /// </summary>
        public int PageNumber { get; set; }

        /// <summary>
        /// 行文本
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// 行索引(从0开始)
        /// </summary>
        public int LineIndex { get; set; }

        /// <summary>
        /// 是否是页眉
        /// </summary>
        public bool IsHeader { get; set; }

        /// <summary>
        /// 是否是页脚
        /// </summary>
        public bool IsFooter
        {
            get
            {
                string pattern = @"^-\d+-$";
                return Regex.IsMatch(Text, pattern);
            }
        }

        public decimal MinX { get; set; }

        public decimal MinY { get; set; }

        /// <summary>
        /// 行文本在word总文本中的下标
        /// </summary>
        public int Offset { get; set; }

        /// <summary>
        /// 长度
        /// </summary>
        public int Length { get; set; }

    }
}
