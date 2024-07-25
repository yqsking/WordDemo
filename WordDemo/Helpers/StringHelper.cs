using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using WordDemo.Helpers;

namespace System
{
    public static class StringHelper
    {

        /// <summary>
        /// 转半角字符
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ConvertCharToHalfWidth(this string str)
        {
            var chatArry = str.ToCharArray();
            // 全角字符到半角字符的Unicode码偏移量
            int offset = 65248;
            StringBuilder stringBuilder = new StringBuilder(chatArry.Length);
            foreach (char c in chatArry)
            {
                // 检查字符是否为全角字符
                if (c >= 65281 && c <= 65374) // 全角字符范围
                {
                    // 转换为半角字符
                    stringBuilder.Append((char)(c - offset));
                }
                else
                {
                    // 如果不是全角字符，直接返回原字符
                    stringBuilder.Append(c);
                }

            }
            return stringBuilder.ToString();
        }

        /// <summary>
        /// 移除空格 下划线 \a \t \r \f \n转义符
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string RemoveSpaceAndEscapeCharacter(this string str)
        {
            return str.Replace("\a", "").Replace("\t", "").Replace("\r", "").Replace("\f", "").Replace(" ", "").Replace("\n", "").Replace("_", "").Replace(":unselected:", "").Replace(":selected:", "");
        }

        /// <summary>
        /// 按字符分割字符串，相邻数字不会被分割
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static List<string> SplitByChar(this string str)
        {
            var strChars = str.ToList().Select(s => new {
                Text = s.ToString(),
                IsNumber = Regex.IsMatch(s.ToString(), @"^-?\d+(\.\d+)?$")
            })
              .ToList();
            var strList = new List<string>();
            var numberList = new List<string>();
            int skipNumber = 0;
            foreach (var strChar in strChars)
            {
                skipNumber++;
                if (strChar.IsNumber)
                {
                    numberList.Add(strChar.Text);
                    var nextStrChar = strChars.Skip(skipNumber).FirstOrDefault();
                    if (nextStrChar == null || !nextStrChar.IsNumber)
                    {
                        strList.Add(string.Join("", numberList));
                        numberList = new List<string>();
                    }
                }
                else
                {
                    strList.Add(strChar.Text);
                }
            }
            return strList;
        }

        /// <summary>
        /// 获取字符串中的日期字符
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetDateString(this string str)
        {
            var dateStringList = new List<string>();
            string datePattern = @"\d{4}年\d{1,2}月\d{1,2}日";
            string yearPattern = @"\d{4}年";
            var matchResult = Regex.Match(str, datePattern);
            if (matchResult.Success)
            {
                return matchResult.Value;
            }
            else
            {
                matchResult = Regex.Match(str, yearPattern);
                return matchResult.Value;
            }

        }

        /// <summary>
        /// 获取多个替换匹配项
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static List<string> GetAllReplaceItemList(this string str)
        {
            var matchReplaceItemList = new List<string>();
            string datePattern = @"\d{4}年\d{1,2}月\d{1,2}日";
            string yearPattern = @"\d{4}年";
            string keyvaluePattern =string.Join("|" ,WordTableConfigHelper.GetCellReplaceItemConfig().SelectMany(s => new[] { s.Key, s.Value }).Distinct().ToList());
            var dateMatchResults= Regex.Matches(str,datePattern);
            foreach(Match dateMatchResult in dateMatchResults)
            {
                matchReplaceItemList.Add(dateMatchResult.Value);
            }
            if(dateMatchResults.Count<=0)
            {
                var yearMatchResults=Regex.Matches(str, yearPattern);
                foreach(Match yearMatchResult in yearMatchResults)
                {
                    matchReplaceItemList.Add(yearMatchResult.Value);
                }
            }

            var keyvalueMatchResults=Regex.Matches(str, keyvaluePattern);
            foreach(Match keyvalueMatchResult in keyvalueMatchResults)
            {
                matchReplaceItemList.Add(keyvalueMatchResult.Value);
            }
            return matchReplaceItemList;
        }

        /// <summary>
        /// 替换文本中的日期为空
        /// </summary>
        /// <param name="str"></param>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static string ReplaceDate(this string str)
        {
            string datePattern = @"\d{4}年\d{1,2}月\d{1,2}日";
            string yearPattern = @"\d{4}年";
            return Regex.Replace(Regex.Replace(str, datePattern, ""), yearPattern, "");
        }

        /// <summary>
        /// 匹配字符串中的word标题，如果不存在返回空字符串
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string MatchWordTitle(this string str)
        {
            string matchTitle = "";
            //标题类型：一、 1. （1） (1) (a)（a）（ii）（ii）
            var titlePatterns = new string[] {
                   @"^[零一二三四五六七八九十]+、", //一、
                   @"^\d+\.",//1.
                   @"^（\d+）",//（1）
                   @"^\(\d+\)",//(1)
                   @"^\d+\)",//1)
                   @"^（[a-z]+）",//（a）
                   @"^\([a-z]+\)",//(a)
                   @"^（[A-Z]+）",//（a）（ii）
                   @"^\([A-Z]+\)",//(a) （ii）
                };
            foreach (var titlePattern in titlePatterns)
            {
                var matchResult = Regex.Match(str, titlePattern);
                if (matchResult.Success)
                {
                    matchTitle = matchResult.Value;
                    break;
                }
            }
            return matchTitle;
        }

        /// <summary>
        /// 移出字符串中的word标题
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string RemoveWordTitle(this string str)
        {
            string title = str.MatchWordTitle();
            return string.IsNullOrWhiteSpace(title) ? str : str.Replace(title, "");
        }

        /// <summary>
        /// 莱文斯坦距离
        /// </summary>
        /// <param name="str1">文本一</param>
        /// <param name="str2">文本二</param>
        /// <returns></returns>
        public static double Levenshtein_Distance(string str1, string str2)
        {
            if (string.IsNullOrEmpty(str1) || string.IsNullOrEmpty(str2))
            {
                return 0d;
            }
            var distance = 0;
            int m = str1.Length;
            int n = str2.Length;
            var maxLength = Math.Max(m, n);
            int[,] lev = new int[m + 1, n + 1];
            // 字符串str1从空串 变为 字符串str2 前j个字符 的莱文斯坦距离
            for (int j = 0; j < n + 1; j++)
            {
                lev[0, j] = j;
            }
            // 字符串str1从前i个字符 变为 空串 的莱文斯坦距离
            for (int i = 1; i < m + 1; i++)
            {
                lev[i, 0] = i;
            }

            for (int i = 1; i < m + 1; i++)
            {
                for (int j = 1; j < n + 1; j++)
                {
                    // 在 字符串A的前i个字符 与 字符串B的前j-1个字符 完全相同的基础上, 进行一次插入操作
                    int countByInsert = lev[i, j - 1] + 1;
                    // 在 字符串A的前i-1个字符 与 字符串B的前j个字符 完全相同的基础上, 进行一次删除操作
                    int countByDel = lev[i - 1, j] + 1;
                    // 在 字符串A的前i-1个字符 与 字符串B的前j-1个字符 完全相同的基础上, 进行一次替换操作
                    int countByReplace = str1[i - 1] == str2[j - 1] ? lev[i - 1, j - 1] : lev[i - 1, j - 1] + 1;
                    // 计算 字符串A的前i个字符 与 字符串B的前j个字符 的莱文斯坦距离
                    lev[i, j] = Math.Min(Math.Min(countByInsert, countByDel), countByReplace);
                }
            }
            distance = lev[m, n];
            var similar = (maxLength - distance) / (maxLength * 1d);
            return similar;//1m - Convert.ToDecimal(distance) / Convert.ToDecimal(maxLength);
        }

        /// <summary>
        /// 替换为空
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        public static string ReplaceEmpty(this string str1,string str2)
        {
            return str1.Replace(str2, "");
        }

        /// <summary>
        /// 是否包含
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        public static bool IsContains(this string str1,string str2)
        {
            return str1.Length>str2.Length?Regex.IsMatch(str1,str2):Regex.IsMatch(str2,str1);
        }

        /// <summary>
        /// 判断集合元素是否连续重复出现，如：年末数 年初数 年末数 年初数
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool HasRepeatedPair(this List<string> list)
        {
            if (list == null || list.Count < 4)
                return false;

            var pairCounts = new Dictionary<Tuple<string, string>, int>();

            for (int i = 0; i < list.Count - 1; i++)
            {
                if (i > 0 && list[i] == list[i - 1] && list[i + 1] == list[i])
                {
                    // 忽略连续的相同元素对
                    continue;
                }

                var pair = Tuple.Create(list[i], list[i + 1]);
                if (pairCounts.ContainsKey(pair))
                {
                    pairCounts[pair]++;
                    if (pairCounts[pair] >= 2)
                    {
                        // 如果元素对出现了两次或以上，返回true
                        return true;
                    }
                }
                else
                {
                    pairCounts.Add(pair, 1);
                }
            }

            return false;
        }

        /// <summary>
        /// 是否word表格数据行（不包含日期，包含任意三位数, 或者任意数字）
        /// </summary>
        /// <param name="rowContent"></param>
        /// <returns></returns>
        public static bool IsWordTableDateRow(this string rowContent)
        {
            bool isContainDate = Regex.IsMatch(rowContent, "(\\d{4}年)|(\\d{1,2})(月|日)");
            bool isContainMoney = Regex.IsMatch(rowContent, "(\\d{3},)|(\\d+)");
            return !isContainDate && isContainMoney;
        }

    }
}
