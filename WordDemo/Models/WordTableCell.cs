using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using WordDemo.Enums;
using WordDemo.Models;

namespace WordDemo
{
    /// <summary>
    /// word表格单元格
    /// </summary>
    public class WordTableCell
    {
        /// <summary>
        /// 原值
        /// </summary>
        public string OldValue {  get; set; }

        /// <summary>
        /// 新值
        /// </summary>
        public string NewValue { get; set; }

        /// <summary>
        /// 操作类型
        /// </summary>
        public OperationTypeEnum OperationType { get; set; }

        /// <summary>
        /// 单元格所在word中的范围
        /// </summary>
        public Range Range { get; set; }
       
        /// <summary>
        /// 单元格开始行索引(从1开始)
        /// </summary>
        public int StartRowIndex {  get; set; }

        /// <summary>
        /// 单元格结束行索引
        /// </summary>
        public int EndRowIndex => RowSpan > 0 ? StartRowIndex + RowSpan - 1 : StartRowIndex;

        /// <summary>
        /// 单元格开始列实际索引(存在列合并的表格，合并后下一列的索引会比实际小)
        /// </summary>
        public int ActualStartColumnIndex {  get; set; }

        /// <summary>
        /// 单元格开始列索引(从1开始)
        /// </summary>
        public int StartColumnIndex { get; set; }

        /// <summary>
        /// 单元格结束列索引
        /// </summary>
        public int EndColumnIndex => ColSpan > 0 ? StartColumnIndex + ColSpan - 1 : StartColumnIndex;

        /// <summary>
        /// 合并行数
        /// </summary>
        public int RowSpan { get; set; } = 1;

        /// <summary>
        /// 用于判断单元格是否垂直合并
        /// </summary>
        public string VMergeVal { get; set; }

        /// <summary>
        /// 合并列数
        /// </summary>
        public int ColSpan { get; set; } = 1;

        /// <summary>
        /// 是否替换单元格
        /// </summary>
        public bool IsReplaceCell { get; set; }

        /// <summary>
        /// 是否表头
        /// </summary>
        public bool IsHeadColumn { get; set; }

        /// <summary>
        /// 最小X轴坐标
        /// </summary>
        public decimal MinX { get; set; }

        /// <summary>
        /// 最小Y轴坐标
        /// </summary>
        public decimal MinY { get; set; }

        /// <summary>
        /// 偏移量
        /// </summary>
        public int Offset { get; set; }

        /// <summary>
        /// 长度
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// 所在页码
        /// </summary>
        public int PageNumber { get; set; }

        /// <summary>
        /// 根据y轴坐标分割的单元格字符
        /// </summary>
        public List<WordChar> CellChars { get; set; }=new List<WordChar>();

        /// <summary>
        /// 根据物理行拆分的子单元格
        /// </summary>
        public List<WordTableCell> ChilderCells=new List<WordTableCell>();

        public decimal YPositiondifference { get; set; }

    }
}
