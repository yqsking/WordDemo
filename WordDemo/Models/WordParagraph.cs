
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WordDemo.Models
{
    /// <summary>
    /// word段落
    /// </summary>
    public class WordParagraph
    {
        /// <summary>
        /// 页码数
        /// </summary>
        public int PageNumber { get; set; }

        /// <summary>
        /// 段落文本
        /// </summary>
        public string Text {  get; set; }

        /// <summary>
        /// 段落索引，从1开始
        /// </summary>
        public int ParagraphNumber { get; set; }

        /// <summary>
        /// 段落Range
        /// </summary>
        public Range Range { get; set; }

        /// <summary>
        /// 段落文本根据制表位分割去除转义符后的结果
        /// </summary>
        public List<string> TabSplitList { 
            get
            {
                if(string.IsNullOrWhiteSpace(Text))
                {
                    return new List<string>();
                }
                return  Text.Split('\t').Select(s => s.RemoveSpaceAndEscapeCharacter()).Where(w => !string.IsNullOrWhiteSpace(w)).ToList();
            }
        }

    }
}
