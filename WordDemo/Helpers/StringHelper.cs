using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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
            var chatArry= str.ToCharArray();
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
            return str.Replace("\a", "").Replace("\t", "").Replace("\r", "").Replace("\f", "").Replace(" ","").Replace("\n","").Replace("_","").Replace(":unselected:", "").Replace(":selected:", "");
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
                IsNumber = Regex.IsMatch(s.ToString(), @"^-?\d+(\.\d+)?$") })
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
            var dateStringList=new List<string>();
            string datePattern = @"\d{4}年\d{1,2}月\d{1,2}日";
            string yearPattern = @"\d{4}年";
            var matchResult=Regex.Match(str, datePattern);
            if(matchResult.Success)
            {
                return matchResult.Value;
            }
            else
            {
                matchResult= Regex.Match(str, yearPattern);
                return matchResult.Value;
            }
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
            return Regex.Replace(Regex.Replace(str, datePattern, ""),yearPattern,"");
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
            foreach(var titlePattern in titlePatterns)
            {
                var matchResult= Regex.Match(str, titlePattern);
                if(matchResult.Success)
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
           string title= str.MatchWordTitle();
            return string.IsNullOrWhiteSpace(title) ? str : str.Replace(title, "");
        }

    }
}
