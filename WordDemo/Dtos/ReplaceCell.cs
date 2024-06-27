using System;
using WordDemo.Enums;

namespace WordDemo.Models
{
    public class ReplaceCell
    {
        /// <summary>
        /// 单元格列索引(水平表头返回的是列索引，垂直表头返回的是行索引)
        /// </summary>
        public int Index { get; set; }
        /// <summary>
        /// 单元格值
        /// </summary>
        public string CellValue { get; set; }
        /// <summary>
        /// 替换匹配项字符
        /// </summary>
        public string ReplaceMatchItem { get; set; }

        /// <summary>
        /// 替换匹配项日期
        /// </summary>
        public DateTime? ReplaceMatchItemDate =>
            !string.IsNullOrWhiteSpace(ReplaceMatchItem) && ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date 
            ? (DateTime?)Convert.ToDateTime(ReplaceMatchItem) : null;

        /// <summary>
        /// 替换匹配项类型
        /// </summary>
        public ReplaceMatchItemTypeEnum ReplaceMatchItemType { get; set; }
    }
}
