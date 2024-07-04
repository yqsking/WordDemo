using System;
using System.Collections.Generic;
using System.Linq;
using WordDemo.Enums;
using WordDemo.Models;

namespace WordDemo
{
    public class WordTable
    {
        /// <summary>
        /// 表格索引，从1开始
        /// </summary>
        public int TableNumber { get; set; }

        /// <summary>
        /// 表格所在页码
        /// </summary>
        public int PageNumber { get; set; }

        /// <summary>
        /// 是否匹配到word段落
        /// </summary>
        public bool IsMatchWordParagraph { get; set; }
      
        /// <summary>
        /// 表格前五条文本
        /// </summary>
        public List<WordLine> UpLines { get; set; }=new List<WordLine>();

        /// <summary>
        /// 过滤页眉页尾的表格前多条文本
        /// </summary>
        public List<WordLine> FilterUpLines => UpLines.Where(w => w != null && !string.IsNullOrWhiteSpace(w.Text) && !w.IsHeader && !w.IsFooter).ToList();

        /// <summary>
        /// 表格后五条文本
        /// </summary>
        public List<WordLine> DownLines { get; set; } = new List<WordLine>();

        /// <summary>
        /// 过滤页眉页尾的表格后多条文本
        /// </summary>
        public List<WordLine> FilterDownLines => DownLines.Where(w => w != null && !string.IsNullOrWhiteSpace(w.Text) && !w.IsHeader && !w.IsFooter).ToList();

        /// <summary>
        /// 表格内容开始段落
        /// </summary>
        public int TableContentStartParagraphNumber {  get; set; }

        /// <summary>
        /// 表格内容结束段落
        /// </summary>
        public int TableContentEndParagraphNumber {  get; set; }

        /// <summary>
        /// 表格在word的段落
        /// </summary>
        public List<WordParagraph> ContentParagraphs { get; set; }= new List<WordParagraph>();

        ///// <summary>
        ///// 是否制表位表格
        ///// </summary>
        //public bool IsTabStopTable => ContentParagraphs.Any() && !ContentParagraphs.All(w => w.Range.Text.Contains("\r\a"));

        /// <summary>
        /// 是否制表位表格
        /// </summary>
        public bool IsTabStopTable => ContentParagraphs.Any() && !ContentParagraphs.All(w => w.OldText.Contains("\r\a"));

        /// <summary>
        /// 表格在word的段落 过滤纯下划线段落
        /// </summary>
        public List<WordParagraph> FilterContentParagraphs => ContentParagraphs.Where(w => !w.Text.ToList().All(text => text.ToString() == "_")).ToList();

        /// <summary>
        /// 表格行
        /// </summary>
        public List<WordTableRow> Rows { get; set; } = new List<WordTableRow>();

        /// <summary>
        /// 表格前三行内容
        /// </summary>
        public List<string> FirstThreeLineTexts => Rows.Where(w => !string.IsNullOrWhiteSpace(w.RowContent))
            .Take(3).Select(s => s.RowContent).ToList();

        /// <summary>
        /// 表格后三行内容
        /// </summary>
        public List<string> LastThreeLineTexts => Rows.Where(w => !string.IsNullOrWhiteSpace(w.RowContent)).Reverse()
            .Take(3).OrderBy(o=>o.RowNumber).Select(s => s.RowContent).ToList();

        /// <summary>
        /// 表头行
        /// </summary>
        public List<WordTableRow> HeadRows => Rows.Where(w => w.IsHeadRow).ToList();

        /// <summary>
        /// 数据行
        /// </summary>
        public List<WordTableRow> DataRows => Rows.Where(w => !w.IsHeadRow).ToList();

        /// <summary>
        /// 表格第一行内容
        /// </summary>
        public string FirstRowContent => Rows.Where(w=>!string.IsNullOrWhiteSpace(w.RowContent.RemoveSpaceAndEscapeCharacter()))
            .FirstOrDefault()?.RowContent;

        /// <summary>
        /// 表格最后一行内容
        /// </summary>
        public string LastRowContent=>Rows.Where(w=>!string.IsNullOrWhiteSpace(w.RowContent.RemoveSpaceAndEscapeCharacter()))
            .LastOrDefault()?.RowContent;

        /// <summary>
        /// 表格首char高度
        /// </summary>
        public decimal FontHeight { get; set; }

        /// <summary>
        /// 操作类型
        /// </summary>
        public OperationTypeEnum OperationType { get; set; } = OperationTypeEnum.NotOperation;

        /// <summary>
        /// 错误消息
        /// </summary>
        public List<string> ErrorMsgs { get; set; }= new List<string>();

    }
}
