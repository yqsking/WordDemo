

namespace WordDemo.Dtos
{
    /// <summary>
    /// 单元格左边距
    /// </summary>
    public class CellLeftMarginInfo
    {
        /// <summary>
        /// 单元格左侧与正文左边距
        /// </summary>
        public float CellMinLeftMargin { get; set; }

        /// <summary>
        /// 单元格右侧与正文左边距
        /// </summary>
        public float CellMaxLeftMargin { get; set; }

        /// <summary>
        /// 单元格内容结束位置(不包含转义符)与正文左边距
        /// </summary>
        public float CellContentEndPointLeftMargin { get; set; }

        /// <summary>
        /// 单元格内容中间位置(不包含转义符)与正文左边距
        /// </summary>
        public float CellContentCenterPointLeftMargin =>(CellMinLeftMargin+CellContentEndPointLeftMargin)/2;

        /// <summary>
        /// 单元格内容小数点与正文左边距
        /// </summary>
        public float CellContentDecimalPointLeftMargin { get; set; }

      




    }
}
