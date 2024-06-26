﻿using System.Collections.Generic;

namespace WordDemo.Helpers
{

    public static class WordTableConfigHelper
    {
        /// <summary>
        /// 获取单元格替换关系配置
        /// </summary>
        /// <returns></returns>
        public static List<List<(string MatchItem,int Sort)>> GetCellReplaceRuleConfig()
        {
            return new Dictionary<string, string>() {
                { "年末数","年初数"},
                { "本年","上年"}
            };
        }

        /// <summary>
        /// 根据表格第一个非空单元格第一个字符最大y轴与最小y轴差获取偏移量
        /// </summary>
        /// <param name="fontHeight">默认0.1</param>
        /// <returns></returns>
        public static decimal GetOffsetValueByFontHeight(decimal fontHeight = 0.1m)
        {

            if (fontHeight >= 0.1m)
            {
                return 0.12m;
            }
            else if (fontHeight < 0.1m && fontHeight >= 0.07m)
            {
                return 0.08m;
            }
            else
            {
                return 0.05m;
            }
        }
    }
}
