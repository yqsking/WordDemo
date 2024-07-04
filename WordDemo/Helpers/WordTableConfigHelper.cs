using System.Collections.Generic;

namespace WordDemo.Helpers
{

    public static class WordTableConfigHelper
    {
        /// <summary>
        /// 获取单元格替换项配置 
        /// </summary>
        /// <returns></returns>
        public static List<KeyValuePair<string,string>> GetCellReplaceItemConfig()
        {
            return new List<KeyValuePair<string, string>> {
               new KeyValuePair<string, string>("本年累计数","上年累计数"),
               new KeyValuePair<string, string>("本年年末数","上年年末数"),
               new KeyValuePair<string, string>("本年期末数","上年期末数"),
               new KeyValuePair<string, string>("本期期末数","上期期末数"),
               new KeyValuePair<string, string>("本期年末数","上期年末数"),
               new KeyValuePair<string, string>("年末数","年初数"),
               new KeyValuePair<string, string>("期末数","期初数"),
               new KeyValuePair<string, string>("本年","上年")
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
