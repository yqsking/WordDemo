using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace WordDemo.Models
{
    public class WordTableRow
    {
        /// <summary>
        /// 行数
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 行单元格
        /// </summary>
        public List<WordTableCell> RowCells { get; set; } = new List<WordTableCell>();

        /// <summary>
        /// 是否表头行
        /// </summary>
        public bool IsHeadRow => RowCells.Any(w => w.IsHeadColumn);

        /// <summary>
        /// 行文本
        /// </summary>
        public string RowContent {
            get {
                return string.Join("", RowCells.Select(s => s.OldValue));
            } 
        }

        /// <summary>
        /// 替换后新行文本
        /// </summary>
        public string NewRowContent
        {
            get {
                return string.Join("", RowCells.Select(s => s.NewValue));
            }
        }

        /// <summary>
        /// 行在word文档中的范围
        /// </summary>
        public Range Range { get; set; }

        /// <summary>
        /// 是否匹配到行段落
        /// </summary>
        public bool IsMatchRowRange { get; set; } = false;
    }
}
