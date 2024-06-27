using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using WordDemo.Enums;
using WordDemo.Helpers;
using WordDemo.Models;
using Table = Microsoft.Office.Interop.Word.Table;

namespace WordDemo
{
    public class WordHelper
    {


        /// <summary>
        /// 获取word制表位表格列表
        /// </summary>
        /// <param name="ocrJson"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static List<WordTable> GetWordTableList(string ocrJson, Document doc)
        {
            var tableList = new List<WordTable>();
            var wordLineList = new List<WordLine>();
            var wordCharList = new List<WordChar>();
            long totalTime = 0;
            var watch = new Stopwatch();

            $"开始解析ocr json文件。。。".Console(ConsoleColor.Yellow);
            #region
            watch.Start();
            JToken docJtoken = JArray.Parse(ocrJson).FirstOrDefault();
            var pageJtokens = docJtoken["pages"].Children();
            foreach (var pageJtoken in pageJtokens)
            {
                int pageNumber = Convert.ToInt32(pageJtoken["page_number"]);
                var wordJtokens = JArray.Parse(pageJtoken["words"].ToString());
                foreach (var wordJtoken in wordJtokens)
                {
                    var polygonJtokens = JArray.Parse(wordJtoken["polygon"].ToString());
                    var positions = polygonJtokens.Select(s => new Position
                    {
                        X = Convert.ToDecimal(s["x"].ToString()),
                        Y = Convert.ToDecimal(s["y"].ToString())
                    });
                    var spanJtoken = wordJtoken["span"];
                    var wordChar = new WordChar
                    {
                        PageNumber = pageNumber,
                        CharNumber = wordCharList.Count,
                        Text = wordJtoken["content"].ToString(),
                        MinX = positions.Min(x => x.X),
                        MinY = positions.Min(y => y.Y),
                        MaxY = positions.Max(y => y.Y),
                        Offset = Convert.ToInt32(spanJtoken["offset"].ToString()),
                        Length = Convert.ToInt32(spanJtoken["length"].ToString())
                    };
                    wordCharList.Add(wordChar);
                }
            }

            wordLineList = GetWordPhysicalLineList(wordCharList, WordTableConfigHelper.GetOffsetValueByFontHeight());

            var tableJtokens = docJtoken["tables"].Children();
            foreach (var tableJtoken in tableJtokens)
            {
                var table = new WordTable();
                try
                {
                    var bounding_regionJtokens = JArray.Parse(tableJtoken["bounding_regions"].ToString());
                    int tableAtPageNumber = Convert.ToInt32(bounding_regionJtokens.FirstOrDefault()["page_number"]);
                    var polygonJtokens = JArray.Parse(bounding_regionJtokens.FirstOrDefault()["polygon"].ToString());
                    var ocrTableRowList = new List<WordTableRow>();
                    var cellJtokens = JArray.Parse(tableJtoken["cells"].ToString());
                    var wordCells = new List<WordTableCell>();
                    foreach (var cell in cellJtokens)
                    {
                        var bounding_region = JArray.Parse(cell["bounding_regions"].ToString()).FirstOrDefault();
                        var cellPolygons = JArray.Parse(bounding_region["polygon"].ToString());
                        int cellAtPageNumber = Convert.ToInt32(bounding_region["page_number"]);
                        var cellPositions = cellPolygons.Select(s => new Position
                        {
                            X = Convert.ToDecimal(s["x"]),
                            Y = Convert.ToDecimal(s["y"])
                        }).ToList();
                        //空字符没有span
                        var span = JArray.Parse(cell["spans"].ToString()).FirstOrDefault();

                        var wordCell = new WordTableCell
                        {
                            // 值为columnHeader时为表头
                            IsHeadColumn = cell["kind"]?.Value<string>() == "columnHeader",
                            StartRowIndex = cell["row_index"].Value<int>() + 1,
                            StartColumnIndex = cell["column_index"].Value<int>() + 1,
                            OldValue = cell["content"].Value<string>().RemoveSpaceAndEscapeCharacter().ConvertCharToHalfWidth(),
                            RowSpan = cell["row_span"].Value<int>(),
                            ColSpan = cell["column_span"].Value<int>(),
                            MinX = cellPositions.Min(m => m.X),
                            MinY = cellPositions.Min(m => m.Y),
                            Offset = span != null ? Convert.ToInt32(span["offset"]) : -1,
                            Length = span != null ? Convert.ToInt32(span["length"]) : -1,
                            YPositiondifference = cellPositions.Max(m => m.Y) - cellPositions.Min(m => m.Y)
                        };
                        wordCells.Add(wordCell);
                    }

                    wordCells.GroupBy(g => g.StartRowIndex).ToList().ForEach(f =>
                    {
                        var tableRow = new WordTableRow()
                        {
                            RowNumber = f.Key,
                            RowCells = f.ToList()
                        };
                        ocrTableRowList.Add(tableRow);
                    });
                    var noEmptyFirstCell = ocrTableRowList.FirstOrDefault().RowCells.Where(w => !string.IsNullOrWhiteSpace(w.OldValue)).FirstOrDefault();
                    var noEmptyFirstCellFirstChar = wordCharList.FirstOrDefault(f => f.Offset == noEmptyFirstCell.Offset);
                    var charHeight = noEmptyFirstCellFirstChar.MaxY - noEmptyFirstCellFirstChar.MinY;

                    //ocr识别数据会出现多个物理行被算成一行 进行物理坐标拆分还原表格
                    var pageWordCharList = wordCharList.Where(w => w.PageNumber == tableAtPageNumber).ToList();
                    var offsetValue = WordTableConfigHelper.GetOffsetValueByFontHeight();
                    ocrTableRowList = GetWordTablePhysicalLineList(ocrTableRowList, pageWordCharList, offsetValue);

                    int columnCount = ocrTableRowList.Max(m => m.RowCells.Count);
                    int rowCount = ocrTableRowList.Count;
                    $"第{tableAtPageNumber}页第个{tableList.Count + 1}表格，{columnCount}列，{rowCount}行,表格文字高度：{charHeight},偏差值：{offsetValue}".Console(ConsoleColor.Blue);

                    table.FontHeight = charHeight;
                    table.TableNumber = tableList.Count + 1;
                    table.PageNumber = tableAtPageNumber;
                    table.Rows = ocrTableRowList;
                    tableList.Add(table);

                }
                catch (Exception ex)
                {
                    $"第{tableList.Count + 1}个表格解析异常，{ex.Message}".Console(ConsoleColor.Red);
                }
            }

            watch.Stop();
            totalTime += watch.ElapsedMilliseconds / 1000;
            #endregion
            $"解析ocr json文件结束,耗时:{watch.ElapsedMilliseconds / 1000}秒。。。".Console(ConsoleColor.Yellow);

            "开始解析word段落".Console(ConsoleColor.Yellow);
            #region
            watch.Restart();
            var paragraphList = new List<WordParagraph>();
            //上一个表格首行内容
            string prevTableFirstRowContent = "";
            //上一个表格尾行内容
            string prevTableLastRowContent = "";
            foreach (Paragraph paragraph in doc.Paragraphs)
            {
                if (string.IsNullOrWhiteSpace(paragraph.Range.Text.RemoveSpaceAndEscapeCharacter()))
                {
                    //跳过空段落
                    continue;
                }
                int wdActiveEndPageNumber = Convert.ToInt32(paragraph.Range.Information[WdInformation.wdActiveEndPageNumber]);
                if (paragraph.Range.Tables.Count > 0)
                {
                    //如果段落中有表格 则表格的非空行算一个段落
                    Table paragraphTable = paragraph.Range.Tables[1];
                    var firstAndLastRowContent = GetTableFirstAndLastContent(paragraphTable);
                    if (firstAndLastRowContent.FirstRowContent != prevTableFirstRowContent
                        && firstAndLastRowContent.LastRowContent != prevTableLastRowContent)
                    {
                        var normalTable = GetWordTable(paragraphTable);
                        foreach (var row in normalTable.Rows)
                        {
                            if (string.IsNullOrWhiteSpace(row.Range.Text.RemoveSpaceAndEscapeCharacter()))
                            {
                                continue;
                            }
                            paragraphList.Add(new WordParagraph
                            {
                                PageNumber = wdActiveEndPageNumber,
                                ParagraphNumber = paragraphList.Count + 1,
                                Text = row.Range.Text.RemoveSpaceAndEscapeCharacter().ConvertCharToHalfWidth(),
                                Range = row.Range
                            });
                        }
                        prevTableFirstRowContent = firstAndLastRowContent.FirstRowContent;
                        prevTableLastRowContent = firstAndLastRowContent.LastRowContent;
                    }

                }
                else
                {
                    var wordParagraph = new WordParagraph()
                    {
                        PageNumber = wdActiveEndPageNumber,
                        ParagraphNumber = paragraphList.Count + 1,
                        Text = paragraph.Range.Text.RemoveSpaceAndEscapeCharacter().ConvertCharToHalfWidth(),
                        Range = paragraph.Range,
                    };
                    paragraphList.Add(wordParagraph);
                }

            }
            watch.Stop();
            totalTime += watch.ElapsedMilliseconds / 10000;
            #endregion
            $"word段落解析结束，耗时{watch.ElapsedMilliseconds / 1000}秒".Console(ConsoleColor.Yellow);

            $"开始匹配OCR表格内容起始段落和单元格Range".Console(ConsoleColor.Yellow);
            #region
            watch.Restart();
            foreach (var table in tableList)
            {
                var tableFirstThreeLineTexts = table.FirstThreeLineTexts.Select(s => s.Replace("-", "").RemoveSpaceAndEscapeCharacter()).ToList();
                var tableLastThreeLineTexts = table.LastThreeLineTexts.Select(s => s.Replace("-", "").RemoveSpaceAndEscapeCharacter()).ToList();

                var rangeParagraphList = paragraphList.Where(w => w.PageNumber == table.PageNumber)
                    .OrderBy(o => o.ParagraphNumber).ToList();
                foreach (var paragraph in rangeParagraphList)
                {
                    //当前段落后三条段落（包含当前段落）
                    var paragraphDownFirstThreelines = paragraphList.Where(w => w.ParagraphNumber >= paragraph.ParagraphNumber)
                        .OrderBy(o => o.ParagraphNumber).Take(3).ToList();
                    var downFirstThreeLineTexts = paragraphDownFirstThreelines.Select(s => s.Text.Replace("-", "")).ToList();

                    //当前段落前三条段落(包含当前段落)
                    var paragraphUpFirstThreeLines = paragraphList.Where(w => w.ParagraphNumber <= paragraph.ParagraphNumber)
                        .OrderByDescending(o => o.ParagraphNumber).Take(3).OrderBy(o => o.ParagraphNumber).ToList();
                    var upFirstThreeLineTexts = paragraphUpFirstThreeLines.Select(s => s.Text.Replace("-", "")).ToList();

                    if (paragraphDownFirstThreelines.Any())
                    {
                        //当前段落后三条段落（包含当前段落）如果包含表格非空前三行数据，当前段落被认定为表格内容开始段落
                        if (tableFirstThreeLineTexts.All(w => downFirstThreeLineTexts.Any(a => a.Contains(w))))
                        {
                            table.TableContentStartParagraphNumber = paragraph.ParagraphNumber;
                        }
                    }
                    if (paragraphUpFirstThreeLines.Any())
                    {
                        //当前段落前三条段落（包含当前段落）如果包含表格非空后三行数据，当前段落被认定为表格内容结束段落
                        if (tableLastThreeLineTexts.All(w => upFirstThreeLineTexts.Any(a => a.Contains(w))))
                        {
                            table.TableContentEndParagraphNumber = paragraph.ParagraphNumber;
                        }

                    }

                    if (table.TableContentStartParagraphNumber > 0 && table.TableContentEndParagraphNumber > 0)
                    {
                        $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent}),表格文字高度：{table.FontHeight},从第{table.TableContentStartParagraphNumber}个段落开始,到第{table.TableContentEndParagraphNumber}个段落结束".Console(ConsoleColor.Yellow);
                        table.IsMatchWordParagraph = true;
                        table.ContentParagraphs = rangeParagraphList.Where(w => w.ParagraphNumber >= table.TableContentStartParagraphNumber &&
                            w.ParagraphNumber <= table.TableContentEndParagraphNumber).ToList();

                        $"表格内容字体大小：{table.ContentParagraphs.FirstOrDefault()?.Range.Font.Size ?? 0}磅".Console(ConsoleColor.Yellow);

                        //匹配单元格Range
                        foreach (WordTableRow row in table.Rows)
                        {
                            Range rowRange = table.ContentParagraphs.Where(w => w.Text.Contains(row.RowContent)).FirstOrDefault()?.Range;
                            if (rowRange == null)
                            {
                                $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})第{row.RowNumber}行({row.RowContent})未能找到匹配段落".Console(ConsoleColor.Red);
                                continue;
                            }
                            row.IsMatchRowRange = true;
                            row.Range = rowRange;
                            Range rowRangeCopy = rowRange.Duplicate;
                            foreach (var cell in row.RowCells)
                            {
                                if (string.IsNullOrWhiteSpace(cell.OldValue))
                                {
                                    continue;
                                }
                                string rowRangeText = rowRangeCopy.Text;
                                Range cellRange = rowRangeCopy.Duplicate;
                                int cellStartIndex = rowRangeText.IndexOf(cell.OldValue);
                                int cellEndIndex = cellStartIndex + cell.OldValue.Length;
                                int moveNumber = rowRangeCopy.Text.Length - cellEndIndex;
                                cellRange.MoveStart(WdUnits.wdCharacter, cellStartIndex);
                                cellRange.MoveEnd(WdUnits.wdCharacter, -moveNumber);
                                cell.Range = cellRange;
                                rowRangeCopy.MoveStart(WdUnits.wdCharacter, cellEndIndex);
                            }
                        }
                        break;
                    }

                }

                if (table.TableContentStartParagraphNumber <= 0 || table.TableContentEndParagraphNumber <= 0)
                {
                    $"第{table.PageNumber}页第{table.TableNumber}个表格未能匹配到Range,表格文本高度：{table.FontHeight}".Console(ConsoleColor.Red);
                    var tableFirstThreeLineTextStr = new StringBuilder();
                    tableFirstThreeLineTexts.ForEach(f =>
                    {
                        tableFirstThreeLineTextStr.AppendLine(f);
                    });
                    $"OCR识别到的表格前三条内容：".Console(ConsoleColor.Yellow);
                    tableFirstThreeLineTextStr.ToString().Console(ConsoleColor.Blue);

                    var tableLastThreeLineTextStr = new StringBuilder();
                    tableLastThreeLineTexts.ForEach(f =>
                    {
                        tableLastThreeLineTextStr.AppendLine(f);
                    });
                    $"OCR识别到的表格后三条内容：".Console(ConsoleColor.Yellow);
                    tableLastThreeLineTextStr.ToString().Console(ConsoleColor.Blue);
                }
            }

            watch.Stop();
            totalTime += watch.ElapsedMilliseconds / 10000;
            #endregion
            int matchSuccessCount = tableList.Count(w => w.TableContentStartParagraphNumber > 0 && w.TableContentEndParagraphNumber > 0);
            $"匹配OCR表格起始段落和单元格Range结束，耗时{watch.ElapsedMilliseconds / 1000}秒".Console(ConsoleColor.Yellow);
            $"总共{tableList.Count}个表格，匹配成功{matchSuccessCount}个".Console(ConsoleColor.Yellow);

           
            $"开始生成单元格新值。。。".Console(ConsoleColor.Yellow);
            #region 
            watch.Restart();
            int lastTableIndex = tableList.IndexOf(tableList.LastOrDefault());
            for (int tableIndex = 0; tableIndex < tableList.Count; tableIndex++)
            {
                string errorMsg = string.Empty;
                var table = tableList[tableIndex];
                try
                {
                    if(!table.IsMatchWordParagraph||table.Rows.Any(w=>!w.IsMatchRowRange))
                    {
                        errorMsg = $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})未能匹配到Word段落范围";
                        table.OperationType = OperationTypeEnum.ChangeColor;
                        table.ErrorMsgs.Add(errorMsg);
                        errorMsg.Console(ConsoleColor.Red);
                        continue;
                    }
                    if (!table.HeadRows.Any())
                    {
                        errorMsg=$"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})未能识别到表头";
                        table.OperationType = OperationTypeEnum.ChangeColor;
                        table.ErrorMsgs.Add(errorMsg);
                        errorMsg.Console(ConsoleColor.Red);
                        continue;
                    }

                    //同表左右替换 判断当前表格所有表头是否包含两个及以上不同日期或者包含任意一组关键字
                    var horizontalHeadRowCells = GetHorizontalMergeTableHeadRow(table.HeadRows);
                    var needReplaceHorizontalHeadCellList = horizontalHeadRowCells.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)).ToList();
                    if (needReplaceHorizontalHeadCellList.Count(w=>w.ReplaceMatchItemType==ReplaceMatchItemTypeEnum.Date)>=2||
                        needReplaceHorizontalHeadCellList.Count(w=>w.ReplaceMatchItemType==ReplaceMatchItemTypeEnum.Keyword)>=2)
                    {
                        //执行同表跨列替换逻辑
                        SameTableCrossColumnReplace(table, needReplaceHorizontalHeadCellList);
                        table.OperationType = OperationTypeEnum.ReplaceText;
                        continue;
                    }

                    //同表上下替换 判断当前表格第一列是否包含两个及以上不同日期或者包含任意一组关键字
                    var verticalHeadRowCells = GetVerticalTableHeadRow(table.Rows);
                    var needReplaceVerticalHeadRowCellList = verticalHeadRowCells.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)).ToList();
                    if(needReplaceVerticalHeadRowCellList.Count(w=>w.ReplaceMatchItemType==ReplaceMatchItemTypeEnum.Date)>=2||
                        needReplaceVerticalHeadRowCellList.Count(w=>w.ReplaceMatchItemType==ReplaceMatchItemTypeEnum.Keyword)>=2)
                    {
                        //执行同表跨行替换逻辑
                        SameTableCrossRowReplace(table, needReplaceVerticalHeadRowCellList);
                        table.OperationType = OperationTypeEnum.ReplaceText;
                        continue;
                    }

                    //跨表上下替换 只有当前表非最后一个表格且与下一个表格表头除了匹配项完全一致
                    if(tableIndex<lastTableIndex)
                    {
                        var nextTable = tableList[tableIndex+1];
                        //判断下一个表是否符合替换条件 
                        if (nextTable.IsMatchWordParagraph && nextTable.Rows.All(w => w.IsMatchRowRange))
                        {
                            //判断当前表格与下一个表格是否表头除开日期部分是否完全一致
                            //且当前表格第一列所有行中内容是否存在任意一项在下一个表格第一列所有行中存在
                            var nextTableHorizontalHeadRowCells = GetHorizontalMergeTableHeadRow(nextTable.HeadRows);
                            string tableHeadRowContent = string.Join("", horizontalHeadRowCells.Select(s => s.CellValue)).ReplaceDate();
                            string nextTableHeadRowContent = string.Join("", nextTableHorizontalHeadRowCells.Select(s => s.CellValue)).ReplaceDate();
                            if (tableHeadRowContent == nextTableHeadRowContent)
                            {
                                //执行跨表替换逻辑
                                CrossTableReplace(table, nextTable);
                                table.OperationType = OperationTypeEnum.ReplaceText;
                                nextTable.OperationType = OperationTypeEnum.ReplaceText;
                                //循环跳过下一个表
                                tableIndex++;
                                
                            }
                        }
                     
                    }
                  
                }
       
                catch(Exception ex)
                {
                    errorMsg = $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})生成单元格新值失败，{ex.Message}";
                    table.OperationType= OperationTypeEnum.ChangeColor;
                    table.ErrorMsgs.Add(errorMsg);
                    errorMsg.Console(ConsoleColor.Red);
                    ex.StackTrace.Console(ConsoleColor.Red);
                }
            }
            watch.Stop();
            totalTime += watch.ElapsedMilliseconds / 10000;
            #endregion
            $"生成单元格新值结束,耗时：{watch.ElapsedMilliseconds / 1000}秒。。。".Console(ConsoleColor.Yellow);

            $"解析表格单元格替换规则总耗时：{totalTime}秒".Console(ConsoleColor.Yellow);

            return tableList;
        }


        #region 正常表格

        /// <summary>
        /// 获取word正常表格
        /// </summary>
        /// <param name="wordTable"></param>
        /// <returns></returns>
        private static WordTable GetWordTable(Table wordTable)
        {
            var table = new WordTable();
            string tableXml = wordTable.Range.WordOpenXML;
            XDocument document = XDocument.Parse(tableXml);
            var rows = document.Root.Descendants().Where(f => f.Name.LocalName == "tr").ToList();
            int rowIndex = 1;
            foreach (var row in rows)
            {
                var rowCells = new List<WordTableCell>();
                var cells = row.Descendants().Where(w => w.Name.LocalName == "tc").ToList();
                int columnIndex = 1;
                foreach (var cell in cells)
                {
                    var tcPr = cell.Descendants().FirstOrDefault(w => w.Name.LocalName == "tcPr");
                    var gridSpan = tcPr.Descendants().FirstOrDefault(w => w.Name.LocalName == "gridSpan");
                    var gridSpanVal = gridSpan == null ? 1 : Convert.ToInt32(gridSpan.Attributes().FirstOrDefault(w => w.Name.LocalName == "val").Value);

                    var vMerge = tcPr.Descendants().FirstOrDefault(w => w.Name.LocalName == "vMerge");
                    var vMergeVal = vMerge == null ? null : vMerge.Attributes().FirstOrDefault(w => w.Name.LocalName == "val")?.Value ?? string.Empty;

                    var tableCell = new WordTableCell()
                    {
                        OldValue = cell.Value,
                        StartRowIndex = rowIndex,
                        ActualStartColumnIndex = cells.IndexOf(cell) + 1,
                        StartColumnIndex = columnIndex,
                        ColSpan = gridSpanVal,
                        VMergeVal = vMergeVal
                    };
                    rowCells.Add(tableCell);
                    if (tableCell.ColSpan > 1)
                    {
                        columnIndex += tableCell.ColSpan;
                    }
                    else
                    {
                        columnIndex++;
                    }
                }

                Microsoft.Office.Interop.Word.Row wordRow = wordTable.Rows[rowIndex];
                var tableRow = new WordTableRow()
                {
                    RowNumber = rowIndex,
                    RowCells = rowCells,
                    Range = wordRow.Range
                };
                table.Rows.Add(tableRow);
                rowIndex++;
            }

            foreach (var row in table.Rows)
            {
                foreach (var cell in row.RowCells)
                {
                    cell.RowSpan = GetVerticalMergeCount(cell, table.Rows);
                    if (cell.VMergeVal != "")
                    {
                        cell.Range = wordTable.Cell(cell.StartRowIndex, cell.ActualStartColumnIndex).Range;
                    }

                }
            }
            //第一个单元格垂直合并数量
            int firstCellRowSpan = table.Rows.FirstOrDefault().RowCells.FirstOrDefault().RowSpan;
            foreach (var row in table.Rows)
            {
                if (row.RowNumber <= firstCellRowSpan)
                {
                    foreach (var cell in row.RowCells)
                    {
                        cell.IsHeadColumn = true;
                    }
                }
                //排除掉被垂直合并的单元格
                row.RowCells = row.RowCells.Where(w => w.VMergeVal != "").ToList();
            }
            return table;
        }

        /// <summary>
        /// 获取表格首行和尾行内容
        /// </summary>
        /// <param name="wordTable"></param>
        /// <returns>FirstRowContent:表格首行内容,LastRowContent:表格尾行内容</returns>
        private static (string FirstRowContent, string LastRowContent) GetTableFirstAndLastContent(Table wordTable)
        {
            string tableXml = wordTable.Range.WordOpenXML;
            XDocument document = XDocument.Parse(tableXml);
            var rows = document.Root.Descendants().Where(f => f.Name.LocalName == "tr").ToList();
            var firstRow = rows.FirstOrDefault();
            var lastRow = rows.LastOrDefault();
            return (firstRow.Value, lastRow.Value);
        }

        /// <summary>
        /// 获取垂直合并单元格数量
        /// </summary>
        /// <param name="currentCell"></param>
        /// <param name="tableRowList"></param>
        /// <returns></returns>
        private static int GetVerticalMergeCount(WordTableCell currentCell, List<WordTableRow> tableRowList)
        {
            int mergeCount = 1;
            if (currentCell.VMergeVal == "restart")
            {
                var cells = tableRowList.SelectMany(s => s.RowCells).Where(w => w.StartRowIndex > currentCell.StartRowIndex && w.StartColumnIndex == currentCell.StartColumnIndex)
                 .OrderBy(o => o.StartRowIndex).ToList();
                if (cells != null && cells.Any())
                {
                    foreach (var cell in cells)
                    {
                        if (cell.VMergeVal == null || cell.VMergeVal == "restart")
                        {
                            break;
                        }
                        mergeCount++;
                    }
                }
            }

            return mergeCount;
        }

        #endregion

        #region 制表位表格

        /// <summary>
        /// 获取word物理行列表
        /// </summary>
        /// <param name="jsonStr">word json文件</param>
        /// <param name="offsetValue">偏移值(两个y轴坐标相减绝对值小于偏移值时，算同一个物理行)</param>
        /// <returns></returns>
        private static List<WordLine> GetWordPhysicalLineList(List<WordChar> wordChars, decimal offsetValue)
        {
            var wordLineList = new List<WordLine>();
            var wordCharPageNumberGroupResults = wordChars.GroupBy(g => g.PageNumber).OrderBy(o => o.Key).ToList();
            foreach (var groupItems in wordCharPageNumberGroupResults)
            {
                var pageWordCharList = groupItems.OrderBy(o => o.Offset).ToList();
                var firstWordChar = pageWordCharList.FirstOrDefault();
                var lineWordChars = new List<WordChar> { firstWordChar };
                foreach (var wordchar in pageWordCharList.Skip(1))
                {
                    if (Math.Abs(firstWordChar.MinY - wordchar.MinY) <= offsetValue)
                    {
                        lineWordChars.Add(wordchar);
                    }
                    else
                    {
                        string lineText = string.Join("", lineWordChars.Select(s => s.Text));
                        wordLineList.Add(new WordLine
                        {
                            PageNumber = firstWordChar.PageNumber,
                            Text = lineText,
                            LineIndex = wordLineList.Count,
                            MinX = firstWordChar.MinX,
                            MinY = firstWordChar.MinY
                        });

                        lineWordChars.Clear();
                        firstWordChar = wordchar;
                        lineWordChars.Add(wordchar);
                    }

                }

                if (lineWordChars.Any())
                {
                    string lineText = string.Join("", lineWordChars.Select(s => s.Text));
                    wordLineList.Add(new WordLine
                    {
                        PageNumber = firstWordChar.PageNumber,
                        Text = lineText,
                        LineIndex = wordLineList.Count,
                        MinX = firstWordChar.MinX,
                        MinY = firstWordChar.MinY
                    });
                }
            }
            return wordLineList;
        }

        /// <summary>
        /// 获取word制表位表格物理行列表
        /// </summary>
        /// <param name="tableRows"></param>
        /// <param name="pageWordChars"></param>
        /// <param name="offsetValue"></param>
        /// <returns></returns>
        private static List<WordTableRow> GetWordTablePhysicalLineList(List<WordTableRow> tableRows, List<WordChar> pageWordChars, decimal offsetValue)
        {
            if (tableRows.Count <= 0)
            {
                return tableRows;
            }

            var wordTableRowList = new List<WordTableRow>();
            var tableCellList = new List<WordTableCell>();
            var offsetValueList = new List<decimal>();

            foreach (var tableRow in tableRows)
            {
                foreach (var rowCell in tableRow.RowCells)
                {
                    //逐行逐个单元格物理拆分 跳过空单元格
                    if (string.IsNullOrWhiteSpace(rowCell.OldValue))
                    {
                        rowCell.ChilderCells.Add(rowCell);
                        continue;
                    }
                    var childerCellList = GetWordTablePhysicalCellList(rowCell, pageWordChars, offsetValue);
                    tableCellList.AddRange(childerCellList);

                }
            }

            tableCellList = tableCellList.OrderBy(o => o.MinY).ToList();
            var firstChilderCell = tableCellList.FirstOrDefault();
            var rowCellList = new List<WordTableCell>() { firstChilderCell };
            foreach (var cell in tableCellList.Skip(1))
            {
                if (Math.Abs(firstChilderCell.MinY - cell.MinY) <= offsetValue)
                {
                    rowCellList.Add(cell);
                }
                else
                {
                    var wordTableRow = new WordTableRow()
                    {
                        RowNumber = wordTableRowList.Count + 1,
                    };
                    rowCellList.OrderBy(o => o.MinX).ToList().ForEach(f =>
                    {
                        f.StartRowIndex = wordTableRowList.Count + 1;
                        wordTableRow.RowCells.Add(f);
                    });

                    wordTableRowList.Add(wordTableRow);
                    rowCellList.Clear();
                    firstChilderCell = cell;
                    rowCellList.Add(cell);
                }

            }
            if (rowCellList.Any())
            {
                var wordTableRow = new WordTableRow()
                {
                    RowNumber = wordTableRowList.Count + 1,
                };
                rowCellList.OrderBy(o => o.MinX).ToList().ForEach(f =>
                {
                    f.StartRowIndex = wordTableRowList.Count + 1;
                    wordTableRow.RowCells.Add(f);
                });
                wordTableRowList.Add(wordTableRow);
            }

            return wordTableRowList;
        }

        /// <summary>
        /// 获取word制表位表格单元格物理单元格列表
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="pageWordChars"></param>
        /// <param name="offsetValue"></param>
        /// <returns></returns>
        private static List<WordTableCell> GetWordTablePhysicalCellList(WordTableCell cell, List<WordChar> pageWordChars, decimal offsetValue)
        {
            var childerCellList = new List<WordTableCell>();
            var yPositionDifferenceList = new List<decimal>();
            if (string.IsNullOrWhiteSpace(cell.OldValue))
            {
                childerCellList.Add(cell);
                return childerCellList;
            }
            //单元格开始偏移量
            int cellStartOffset = cell.Offset;
            //单元格结束偏移量
            int cellEndOffset = cell.Offset + cell.Length - 1;

            //当前单元格偏移量范围所包含的所有字符
            var cellWordChars = pageWordChars.Where(w => w.Offset >= cellStartOffset && w.Offset <= cellEndOffset)
                .OrderBy(o => o.Offset).ToList();
            string joinCharStr = string.Join("", cellWordChars.Select(s => s.Text));
            if (cell.OldValue == joinCharStr)
            {
                var firstCellWordChar = cellWordChars.FirstOrDefault();
                //处于同物理行的字符集合
                var lineWordChars = new List<WordChar>() { firstCellWordChar };
                foreach (var cellWordChar in cellWordChars.Skip(1))
                {
                    var yPositionDifference = Math.Abs(firstCellWordChar.MinY - cellWordChar.MinY);
                    yPositionDifferenceList.Add(yPositionDifference);
                    //最小y轴坐标差小于偏差值 算同行
                    if (yPositionDifference <= offsetValue)
                    {
                        lineWordChars.Add(cellWordChar);
                    }
                    else
                    {
                        string lineText = string.Join("", lineWordChars.Select(s => s.Text));
                        childerCellList.Add(new WordTableCell
                        {
                            PageNumber = firstCellWordChar.PageNumber,
                            OldValue = lineText,
                            StartRowIndex = cell.StartRowIndex + childerCellList.Count,
                            StartColumnIndex = cell.StartColumnIndex,
                            MinX = firstCellWordChar.MinX,
                            MinY = firstCellWordChar.MinY,
                            Offset = firstCellWordChar.Offset,
                            Length = lineText.Length,
                            IsHeadColumn = cell.IsHeadColumn,
                            ColSpan = cell.ColSpan
                        });
                        lineWordChars.Clear();
                        firstCellWordChar = cellWordChar;
                        lineWordChars.Add(cellWordChar);
                    }
                }
                if (lineWordChars.Any())
                {
                    string lineText = string.Join("", lineWordChars.Select(s => s.Text));
                    childerCellList.Add(new WordTableCell
                    {
                        PageNumber = firstCellWordChar.PageNumber,
                        OldValue = lineText,
                        StartRowIndex = cell.StartRowIndex + childerCellList.Count,
                        StartColumnIndex = cell.StartColumnIndex,
                        MinX = firstCellWordChar.MinX,
                        MinY = firstCellWordChar.MinY,
                        Offset = firstCellWordChar.Offset,
                        Length = lineText.Length,
                        IsHeadColumn = cell.IsHeadColumn,
                        ColSpan = cell.ColSpan
                    });
                }
            }
            else
            {
                //拆分失败 返回原来的单元格
                childerCellList.Add(cell);
            }

            return childerCellList;
        }

        /// <summary>
        /// 获取行合并水平方向表头
        /// </summary>
        /// <returns>ColumnIndex:列索引,从1开始;CellValue:单元格值;ReplaceMatchItem:替换匹配项,不为空代表当前单元格属于需要替换数据的表头;ReplaceMatchItemType:替换匹配项类型(日期/关键字)</returns>
        private static List<ReplaceCell> GetHorizontalMergeTableHeadRow(List<WordTableRow> wordTableHeadRows)
        {
            var mergeHeadRowCells = new List<ReplaceCell>();
            //补全水平合并单元格表头
            var completionWordTableHeadRows = new List<WordTableRow>();
            foreach (var row in wordTableHeadRows)
            {
                var rowCellList = new List<WordTableCell>();
                foreach (var cell in row.RowCells)
                {
                    for (int i = 1; i <= cell.ColSpan; i++)
                    {
                        var tempCell = new WordTableCell()
                        {
                            PageNumber = cell.PageNumber,
                            OldValue = cell.OldValue,
                            StartColumnIndex = cell.StartColumnIndex + i - 1,
                            ColSpan = 1,
                            StartRowIndex = cell.StartRowIndex,
                            RowSpan = cell.RowSpan
                        };
                        rowCellList.Add(tempCell);
                    }
                }
                completionWordTableHeadRows.Add(new WordTableRow { RowNumber = row.RowNumber, RowCells = rowCellList });
            }
            
            //同列表头内容合并 
            var headCellList = completionWordTableHeadRows.SelectMany(s => s.RowCells).ToList();
            int maxStartColumnIndex = headCellList.Max(m => m.StartColumnIndex);
            for (int i = 1; i <= maxStartColumnIndex; i++)
            {
                var columnCellList = headCellList.Where(w => w.StartColumnIndex == i).OrderBy(o => o.StartRowIndex)
                    .Select(s => s.OldValue).ToList();
                string columnCellJoinValue = string.Join("", columnCellList);
                var getCellContainResult = GetCellContainReplaceMatchItem(columnCellJoinValue);
                mergeHeadRowCells.Add(
                    new ReplaceCell
                    {
                        Index = i,
                        CellValue = columnCellJoinValue,
                        ReplaceMatchItem = getCellContainResult.ReplaceMathItem,
                        ReplaceMatchItemType = getCellContainResult.ReplaceMatchItemType
                    });
            }
            return mergeHeadRowCells;
        }

        /// <summary>
        /// 获取垂直方向表头
        /// </summary>
        /// <param name="wordTableRows"></param>
        /// <returns>RowIndex:行索引,从1开始;CellValue:单元格值;ReplaceMatchItem:替换匹配项,不为空代表当前单元格属于需要替换数据的表头</returns>
        private static List<ReplaceCell> GetVerticalTableHeadRow(List<WordTableRow> wordTableRows)
        {
            var headRowCells = new List<ReplaceCell>();
            var firstColumnCellList = wordTableRows.SelectMany(s => s.RowCells).Where(w => w.StartColumnIndex == 1)
                .OrderBy(o => o.StartRowIndex).ToList();
            foreach (var cell in firstColumnCellList)
            {
                var getCellContainResult = GetCellContainReplaceMatchItem(cell.OldValue);
                headRowCells.Add(new ReplaceCell
                {
                    Index = cell.StartRowIndex,
                    CellValue = cell.OldValue,
                    ReplaceMatchItem = getCellContainResult.ReplaceMathItem,
                    ReplaceMatchItemType = getCellContainResult.ReplaceMatchItemType
                });
            };
            return headRowCells;
        }

        /// <summary>
        /// 获取单元格包含的替换匹配项
        /// </summary>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private static (string ReplaceMathItem, ReplaceMatchItemTypeEnum ReplaceMatchItemType) GetCellContainReplaceMatchItem(string cellValue)
        {
            string replaceMatchItem = cellValue.GetDateString();
            ReplaceMatchItemTypeEnum replaceMatchItemType = ReplaceMatchItemTypeEnum.Date;
            if (string.IsNullOrWhiteSpace(replaceMatchItem))
            {
                //包含匹配键值对
                var replaceMatchItemList = WordTableConfigHelper.GetCellReplaceItemConfig().Select(s => new string[] { s.Key, s.Value }).SelectMany(s => s).ToList();
                replaceMatchItem = replaceMatchItemList.FirstOrDefault(matchItem => cellValue.Contains(matchItem));
                replaceMatchItemType = ReplaceMatchItemTypeEnum.Keyword;
            }
            return (replaceMatchItem, replaceMatchItemType);
        }

        /// <summary>
        /// 同表跨行替换
        /// </summary>
        /// <param name="table"></param>
        /// <param name="replaceCells"></param>
        private static void SameTableCrossColumnReplace(WordTable table, List<ReplaceCell> replaceCells)
        {
            var allCellList = table.Rows.SelectMany(s => s.RowCells).ToList();
            if (replaceCells.All(w => w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date))
            {
                //如果匹配到的是日期 要替换的表头移除掉匹配项文本 得到的表头单元格值应该是一致的
                replaceCells.ForEach(f => f.CellValue = f.CellValue.Replace(f.ReplaceMatchItem, ""));
                //根据表头新单元格值分组
                var headCellValueGroupbyResults = replaceCells.GroupBy(g => g.CellValue).ToList();
                foreach (var headCellValueGroupbyResult in headCellValueGroupbyResults)
                {
                    if (headCellValueGroupbyResult.Count() <= 1)
                    {
                        continue;
                    }
                    //所有替换数据匹配列根据匹配日期降序排序 
                    var replaceCellGroupResults = headCellValueGroupbyResult.OrderByDescending(o => o.ReplaceMatchItemDate).ToList();
                    var maxDateReplaceCell = replaceCellGroupResults.FirstOrDefault();
                    //判断当前表格是季度表还是年度表
                    var yearGroupResults = replaceCellGroupResults.GroupBy(g => g.ReplaceMatchItemDate.Value.Year).ToList();
                    bool isQuarterTable = yearGroupResults.Any(w => w.Count() > 1);
                    //更新第一个表格匹配列
                    allCellList.Where(w => w.StartColumnIndex == maxDateReplaceCell.Index).ToList()
                        .ForEach(cell =>
                        {
                            if (cell.IsHeadColumn)
                            {
                                cell.NewValue = GetNextMaxDateHeadCellValue(isQuarterTable, maxDateReplaceCell.ReplaceMatchItemDate.Value, cell.OldValue);
                            }
                            else
                            {
                                cell.NewValue = "-";
                            }
                            cell.IsReplaceValue = true;
                        });

                    //其他列依次从左往右平移
                    for (int i = 1; i < replaceCellGroupResults.Count; i++)
                    {
                        var currentReplaceHeadCell = replaceCellGroupResults[i];
                        var dataSourceReplaceHeadCell = replaceCellGroupResults[i - 1];
                        //前一列所有单元格
                        var dataSourceCellList = allCellList.Where(w => w.StartColumnIndex == dataSourceReplaceHeadCell.Index).ToList();
                        //当前列所有单元格
                        allCellList.Where(w => w.StartColumnIndex == currentReplaceHeadCell.Index).ToList()
                        .ForEach(cell =>
                        {
                            cell.NewValue = dataSourceCellList.FirstOrDefault(w => w.StartRowIndex == cell.StartRowIndex)?.OldValue;
                            cell.IsReplaceValue = true;
                        });
                    }

                }
            }
            else
            {
                //如果匹配到的是关键字
                var replaceItemList = WordTableConfigHelper.GetCellReplaceItemConfig();
                var alreadyReplaceMatchItems = new List<string>();
                foreach (var replaceCell in replaceCells)
                {
                    var replaceItem = replaceItemList.FirstOrDefault(w => w.Key == replaceCell.ReplaceMatchItem || w.Value == replaceCell.ReplaceMatchItem);
                    var keyReplaceCell = replaceCells.FirstOrDefault(w => w.ReplaceMatchItem == replaceItem.Key);
                    var valueReplaceCell = replaceCells.FirstOrDefault(w => w.ReplaceMatchItem == replaceItem.Value);
                    if (keyReplaceCell != null && valueReplaceCell != null)
                    {
                        if (!alreadyReplaceMatchItems.Contains(keyReplaceCell.ReplaceMatchItem + "_" + valueReplaceCell.ReplaceMatchItem))
                        {
                            foreach (var cell in allCellList.Where(w => w.StartColumnIndex == keyReplaceCell.Index && !w.IsHeadColumn))
                            {
                                cell.NewValue = "-";
                                cell.IsReplaceValue = true;
                            }

                            foreach (var cell in allCellList.Where(w => w.StartColumnIndex == valueReplaceCell.Index && !w.IsHeadColumn))
                            {
                                var dataSourceCell = allCellList.FirstOrDefault(w => w.StartColumnIndex == keyReplaceCell.Index && w.StartRowIndex == cell.StartRowIndex);
                                cell.NewValue = dataSourceCell?.OldValue;
                                cell.IsReplaceValue = true;
                            }

                            alreadyReplaceMatchItems.Add(keyReplaceCell.ReplaceMatchItem + "_" + valueReplaceCell.ReplaceMatchItem);
                        }

                    }
                }
            }
        }

        /// <summary>
        /// 同表跨行替换
        /// </summary>
        /// <param name="table"></param>
        /// <param name="replaceCells"></param>
        private static void SameTableCrossRowReplace(WordTable table, List<ReplaceCell> replaceCells)
        {
            var allCellList = table.Rows.SelectMany(s => s.RowCells).ToList();
            if (replaceCells.All(w => w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date))
            {

            }
            else
            {

            }
        }

        /// <summary>
        /// 跨表替换
        /// </summary>
        /// <param name="table"></param>
        /// <param name="nextWordTable"></param>
        /// <param name="replaceCells"></param>
        private static void CrossTableReplace(WordTable table, WordTable nextWordTable)
        {

        }

        /// <summary>
        /// 获取下一个日期
        /// </summary>
        /// <param name="isQuarterTable"></param>
        /// <param name="date"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private static string GetNextMaxDateHeadCellValue(bool isQuarterTable, DateTime date, string cellValue)
        {
            DateTime? nextMaxDate = null;
            if (isQuarterTable)
            {
                if (date.Month == 1)
                {
                    nextMaxDate = new DateTime(date.Year, 6, 30);
                }
                else if (date.Month == 6)
                {
                    nextMaxDate = new DateTime(date.Year, 12, 31);
                }
                else
                {
                    nextMaxDate = new DateTime(date.Year + 1, 1, 1);
                }
            }
            else
            {
                nextMaxDate = new DateTime(date.Year + 1, date.Month, date.Day);
            }

            cellValue = Regex.Replace(cellValue, @"\d{4}年", nextMaxDate.Value.Year + "年");
            cellValue = Regex.Replace(cellValue, @"\d{1,2}月", nextMaxDate.Value.Month + "月");
            cellValue = Regex.Replace(cellValue, @"\d{1,2}日", nextMaxDate.Value.Day + "日");
            return cellValue;

        }
        #endregion


    }
}


