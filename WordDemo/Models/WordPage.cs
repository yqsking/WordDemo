using System.Collections.Generic;

namespace WordDemo.Models
{
    /// <summary>
    /// word页面
    /// </summary>
    public class WordPage
    {
        /// <summary>
        /// 页码
        /// </summary>

        public int PageNumber { get; set; }

        /// <summary>
        /// 当前页所有页眉文本
        /// </summary>
        public List<WordPageHeader> WordPageHeaders { get; set; } = new List<WordPageHeader>();

        /// <summary>
        /// 当前页所有段落
        /// </summary>
        public List<WordParagraph> WordParagraphs { get; set; } = new List<WordParagraph>();

        /// <summary>
        /// 当前页所有行文本
        /// </summary>
        public List<WordLine> WordLines { get; set; }= new List<WordLine>();

        /// <summary>
        /// 当前页所有字符
        /// </summary>
        public List<WordChar> WordChars { get; set; }=new List<WordChar>();

        /// <summary>
        /// 当前页所有表格
        /// </summary>
        public List<WordTable> WordTables { get; set; } = new List<WordTable>();


    }
}
