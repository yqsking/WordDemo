using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using WordDemo.Dtos;
using WordDemo.Enums;
using WordDemo.Helpers;
using WordDemo.Models;
using Table = Microsoft.Office.Interop.Word.Table;

namespace WordDemo
{
    public class WordHelper
    {
        /// <summary>
        /// 格式化表格表头和添加下划线
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="config"></param>
        /// <returns></returns>
        public static void FormattingWordTableHeaderAndAppendUnderline(Document doc, FormattingWordTableConfig config)
        {
            var wordTableList = new List<WordTable>();
            int tableNumber = 0;
            try
            {
                foreach (Table table in doc.Tables)
                {
                    tableNumber++;
                    Cell firstCell = table.Cell(1, 1);
                    Cell lastCell = table.Range.Cells[table.Range.Cells.Count];
                    $"第{tableNumber}个表格第1个单元格值：{firstCell.Range.Text.RemoveSpaceAndEscapeCharacter()},最后一个单元格值：{lastCell.Range.Text.RemoveSpaceAndEscapeCharacter()}".Console();

                    var wordTable = GetWordTable(table);
                    int headRowStartIndex = wordTable.HeadRows.Min(m => m.RowNumber);
                    int headRowEndIndex = wordTable.HeadRows.Max(m => m.RowNumber);

                    var firstCellBorderList = new List<Border> {
                       firstCell.Range.Borders[WdBorderType.wdBorderTop],
                       firstCell.Range.Borders[WdBorderType.wdBorderLeft],
                       firstCell.Range.Borders[WdBorderType.wdBorderRight]
                    };
                    var lastCellBorderList = new List<Border> {
                        lastCell.Range.Borders[WdBorderType.wdBorderTop],
                        lastCell.Range.Borders[WdBorderType.wdBorderLeft],
                        lastCell.Range.Borders[WdBorderType.wdBorderRight]
                    };
                    var solidLineBorderList = new WdLineStyle[] {
                      WdLineStyle.wdLineStyleSingle,//单实线
                      WdLineStyle.wdLineStyleDouble,//双实线
                      WdLineStyle.wdLineStyleTriple,//三条细实线
                      WdLineStyle.wdLineStyleThinThickSmallGap,WdLineStyle.wdLineStyleThickThinSmallGap,
                      WdLineStyle.wdLineStyleThinThickThinSmallGap,WdLineStyle.wdLineStyleThinThickMedGap,
                      WdLineStyle.wdLineStyleThickThinMedGap,WdLineStyle.wdLineStyleThinThickThinMedGap,
                      WdLineStyle.wdLineStyleThinThickLargeGap,WdLineStyle.wdLineStyleThickThinLargeGap,
                      WdLineStyle.wdLineStyleThinThickThinLargeGap,
                      WdLineStyle.wdLineStyleSingleWavy,//波浪单实线
                      WdLineStyle.wdLineStyleDoubleWavy,//波浪双实线
                    };

                    //如果第一个单元格和最后一个单元格任意一个的上左右边框线都是实线 代表表格是实线
                    if (firstCellBorderList.All(w => solidLineBorderList.Contains(w.LineStyle)) ||
                       lastCellBorderList.All(w => solidLineBorderList.Contains(w.LineStyle)))
                    {

                    }
                    else
                    {
                        //按照虚线表格处理
                    }
                }

            }
            catch (Exception ex)
            {
                $"格式化异常,{ex.Message}".Console(ConsoleColor.Red);
            }

            //cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        /// <summary>
        /// 获取word制表位表格列表
        /// </summary>
        /// <param name="ocrJson">ocr识别的json文件</param>
        /// <param name="doc">word文档对象</param>
        /// <returns></returns>
        public static List<WordTable> GetWordTableList(string ocrJson, Document doc, CancellationToken Cancel = default, Action<object,Events.NodeNotifyEventArgs> errorMsg_Event = null)
        {
            var tableList = new List<WordTable>();
            var normalTableList = new List<WordTable>();
            var paragraphList = new List<WordParagraph>();
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
                    Cancel.ThrowIfCancellationRequested();
                }
            }

            wordLineList = GetWordPhysicalLineList(wordCharList, WordTableConfigHelper.GetOffsetValueByFontHeight());

            var tableJtokens = docJtoken["tables"].Children();
            foreach (var tableJtoken in tableJtokens)
            {
                Cancel.ThrowIfCancellationRequested();
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
                        Cancel.ThrowIfCancellationRequested();
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
            Paragraph while_paragraph = doc.Paragraphs.First;
            while (while_paragraph != null)
            {
                int wdActiveEndPageNumber = Convert.ToInt32(while_paragraph.Range.Information[WdInformation.wdActiveEndPageNumber]);
                $"正在解析第{wdActiveEndPageNumber}页段落：{while_paragraph.Range.Text.RemoveSpaceAndEscapeCharacter()}".Console(ConsoleColor.Yellow);
                if (while_paragraph.Range.Tables.Count > 0)
                {
                    //如果段落中有表格 则表格的非空行算一个段落
                    Table paragraphTable = while_paragraph.Range.Tables[1];
                    var firstAndLastRowContent = GetTableFirstAndLastContent(paragraphTable);
                    var normalTable = GetWordTable(paragraphTable);
                    if (normalTable == null)
                    {
                        while_paragraph = paragraphTable.Range.Cells[paragraphTable.Range.Cells.Count].Range.Paragraphs.Last;
                        while_paragraph = while_paragraph.Next(2);
                        continue;
                    }
                    int tableContentStartParagraphNumber = paragraphList.Count() + 1;
                    foreach (var row in normalTable.Rows)
                    {
                        if (string.IsNullOrWhiteSpace(row.RowContent.RemoveSpaceAndEscapeCharacter()))
                        {
                            //下一个段落
                            while_paragraph = while_paragraph.Next();
                            continue;
                        }
                        paragraphList.Add(new WordParagraph
                        {
                            PageNumber = wdActiveEndPageNumber,
                            ParagraphNumber = paragraphList.Count + 1,
                            OldText = row.Range?.Text ?? "",
                            Text = row.RowContent.RemoveSpaceAndEscapeCharacter().ConvertCharToHalfWidth(),
                            Range = row.Range,
                            IsUsed = true
                        });
                    }
                    int tableContentEndParagraphNumber = paragraphList.Count;
                    normalTable.TableSourceType = TableSourceTypeEnum.WordTable;
                    normalTable.TableNumber = normalTableList.Count + 1;
                    normalTable.PageNumber = wdActiveEndPageNumber;
                    normalTable.TableContentStartParagraphNumber = tableContentStartParagraphNumber;
                    normalTable.TableContentEndParagraphNumber = tableContentEndParagraphNumber;
                    normalTableList.Add(normalTable);
                    while_paragraph = paragraphTable.Range.Cells[paragraphTable.Range.Cells.Count].Range.Paragraphs.Last.Next();
                }
                else
                {
                    string paragraphText = while_paragraph.Range.Text;
                    var wordParagraph = new WordParagraph()
                    {
                        PageNumber = wdActiveEndPageNumber,
                        ParagraphNumber = paragraphList.Count + 1,
                        OldText = paragraphText,
                        Text = paragraphText.RemoveSpaceAndEscapeCharacter().ConvertCharToHalfWidth(),
                        Range = while_paragraph.Range,
                    };
                    paragraphList.Add(wordParagraph);
                }
                //下一个段落
                while_paragraph = while_paragraph.Next();
                Cancel.ThrowIfCancellationRequested();
            }
            watch.Stop();
            totalTime += watch.ElapsedMilliseconds / 10000;
            #endregion
            $"word段落解析结束，耗时{watch.ElapsedMilliseconds / 1000}秒".Console(ConsoleColor.Yellow);

            $"开始匹配OCR表格内容起始段落和单元格Range".Console(ConsoleColor.Yellow);
            #region
            watch.Restart();
            //foreach (var table in tableList)
            var tbCount = tableList.Count;
            for (int i = 0; i < tbCount; i++)
            {
                Cancel.ThrowIfCancellationRequested();
                var table = tableList[i];

                var tableFirstThreeLineTexts = table.FirstThreeLineTexts.Select(s => s.Replace("-", "").RemoveSpaceAndEscapeCharacter().RemoveWordTitle()).ToList();
                var tableLastThreeLineTexts = table.LastThreeLineTexts.Select(s => s.Replace("-", "").RemoveSpaceAndEscapeCharacter().RemoveWordTitle()).ToList();

                var rangeParagraphList = paragraphList.Where(w => w.PageNumber == table.PageNumber && !w.IsEmptyParagraph)
                    .OrderBy(o => o.ParagraphNumber).ToList();
                foreach (var paragraph in rangeParagraphList)
                {
                    //当前段落后三条段落（包含当前段落）
                    var paragraphDownFirstThreelines = paragraphList.Where(w => !w.IsEmptyParagraph && w.ParagraphNumber >= paragraph.ParagraphNumber)
                        .OrderBy(o => o.ParagraphNumber).Take(3).ToList();
                    var downFirstThreeLineTexts = paragraphDownFirstThreelines.Select(s => s.Text.Replace("-", "").RemoveWordTitle()).ToList();

                    //当前段落前三条段落(包含当前段落)
                    var paragraphUpFirstThreeLines = paragraphList.Where(w => !w.IsEmptyParagraph && w.ParagraphNumber <= paragraph.ParagraphNumber)
                        .OrderByDescending(o => o.ParagraphNumber).Take(3).OrderBy(o => o.ParagraphNumber).ToList();
                    var upFirstThreeLineTexts = paragraphUpFirstThreeLines.Select(s => s.Text.Replace("-", "").RemoveWordTitle()).ToList();

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

                        var topFiveParagraphs = paragraphList.Where(w => w.ParagraphNumber < table.TableContentStartParagraphNumber)
                          .OrderByDescending(o => o.ParagraphNumber).Take(5).OrderBy(o => o.ParagraphNumber).ToList();

                        //如果表头为空，重新计算表头 ;如果数据行单元格不全，补充空单元格
                        SupplementCell(table, topFiveParagraphs);

                        table.IsMatchWordParagraph = true;
                        var tableRangeParagraphList = rangeParagraphList.Where(w => w.ParagraphNumber >= table.TableContentStartParagraphNumber &&
                            w.ParagraphNumber <= table.TableContentEndParagraphNumber).ToList();

                        //lxz 2024-07-25 判断ocr 识别表格把表格最后一个单元格识别为下一行，则进行修正；上海汇众汽车车桥系统有限公司.docx 货币资金
                        var tableCount = table.Rows.Count;
                        if (tableRangeParagraphList.Count < tableCount)
                        {

                            var remove_tbIndexList = new List<int>();

                            var tbRowIndex = 0;
                            for (int _i = 0; _i < tableRangeParagraphList.Count; _i++)
                            {
                                var t_para = tableRangeParagraphList[_i];
                                if (tableRangeParagraphList.Any(x => x.OldText.Contains("\r\a")))
                                {
                                    break;
                                }

                                if (string.IsNullOrWhiteSpace(t_para.OldText))
                                {
                                    tbRowIndex++;
                                    continue;
                                }
                                if (tableCount <= tbRowIndex)
                                {
                                    break;
                                }
                                var t_row = table.Rows[tbRowIndex];
                                if (!string.IsNullOrWhiteSpace(t_row.RowContent.Trim()) && tbRowIndex + 1 < tableCount)
                                {
                                    var next_t_Row = table.Rows[tbRowIndex + 1];

                                    if (!string.IsNullOrWhiteSpace(next_t_Row.RowContent) && t_para.Text.Contains($"{t_row.RowContent.Trim()}{next_t_Row.RowContent.Trim()}"))
                                    {
                                        var cells = next_t_Row.RowCells.Where(x => !string.IsNullOrWhiteSpace(x.OldValue));

                                        var t_row_cells_count = t_row.RowCells.Count;
                                        var splists = t_para.OldText.Split('\t');

                                        if (t_row_cells_count > 0 && cells.Count() > 0 && splists != null && splists.Length > 0 && splists.Length == t_row_cells_count + cells.Count())
                                        {
                                            var colIndex = t_row_cells_count;
                                            foreach (var item in cells)
                                            {
                                                colIndex++;
                                                item.StartRowIndex -= 1;
                                                item.StartColumnIndex = colIndex;
                                                t_row.RowCells.Add(item);
                                            }
                                            remove_tbIndexList.Add(tbRowIndex + 1);
                                        }
                                        tbRowIndex++;
                                    }
                                }
                                tbRowIndex++;
                            }
                            if (remove_tbIndexList.Count > 0)
                            {
                                remove_tbIndexList.Reverse();
                                foreach (var item in remove_tbIndexList)
                                {
                                    table.Rows.RemoveAt(item);
                                }

                                for (int _r = 0; _r < table.Rows.Count; _r++)
                                {
                                    table.Rows[_r].RowNumber = _r + 1;
                                    table.Rows[_r].RowCells.ForEach(x =>
                                    {
                                        x.StartRowIndex = _r + 1;
                                    });
                                }
                            }
                        }


                        foreach (var tableRangeParagraph in tableRangeParagraphList)
                        {
                            tableRangeParagraph.IsUsed = true;
                        }
                        table.ContentParagraphs = tableRangeParagraphList;

                        //只计算制表位表格单元格Range
                        if (table.IsTabStopTable)
                        {
                            MatchTabStopTableCellRange(table);
                            if (table.OperationType == OperationTypeEnum.ConsoleError)
                            {
                                //计算表格单元格Range失败 所有段落恢复成未使用
                                tableRangeParagraphList.ForEach(f => {
                                    f.IsUsed = false;
                                });
                            }
                        }
                        break;
                    }

                }

                if (table.TableContentStartParagraphNumber <= 0 || table.TableContentEndParagraphNumber <= 0)
                {
                    var errorMsg = new StringBuilder();
                    errorMsg.AppendLine($"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})未能匹配到Word段落!");
                    errorMsg.AppendLine("OCR识别到的表格前三条内容：");
                    tableFirstThreeLineTexts.ForEach(f =>
                    {
                        errorMsg.AppendLine(f);
                    });
                    errorMsg.AppendLine("OCR识别到的表格后三条内容：");
                    tableLastThreeLineTexts.ForEach(f =>
                    {
                        errorMsg.AppendLine(f);
                    });
                    table.OperationType = OperationTypeEnum.ConsoleError;
                    table.ErrorMsgs.Add(errorMsg.ToString());
                    errorMsg.ToString().Console(ConsoleColor.Red);
                }
            }

            watch.Stop();
            totalTime += watch.ElapsedMilliseconds / 10000;
            #endregion
            int matchSuccessCount = tableList.Count(w => w.TableContentStartParagraphNumber > 0 && w.TableContentEndParagraphNumber > 0);
            $"匹配OCR表格起始段落和单元格Range结束，耗时{watch.ElapsedMilliseconds / 1000}秒".Console(ConsoleColor.Yellow);
            $"总共{tableList.Count}个表格，匹配成功{matchSuccessCount}个;有{tableList.Count(w => w.IsTabStopTable)}个制表位表格,{normalTableList.Count}个正常表格".Console(ConsoleColor.Yellow);

            $"开始生成单元格新值。。。".Console(ConsoleColor.Yellow);
            #region 
            watch.Restart();

            //根据没有使用的段落 还原制表位表格
            var notUseParagraphList = paragraphList.Where(w => !w.IsUsed).ToList();
            var identifyFailTabStopTableList = GetIdentifyFailTabStopTableList(notUseParagraphList, tableList);
            if (identifyFailTabStopTableList.Any())
            {
                foreach (var table in identifyFailTabStopTableList)
                {
                    table.TableNumber = tableList.Count + 1;
                    tableList.Add(table);
                }
            }

            Cancel.ThrowIfCancellationRequested();
            SplitTables(tableList, errorMsg_Event);


            Cancel.ThrowIfCancellationRequested();
            //生成制表位单元格新值
            BuildTabStopTableCellNewValue(tableList, errorMsg_Event);

            //生成正常表格单元格新值
            if (normalTableList.Any())
            {
                MergeTables(normalTableList, errorMsg_Event);
                SplitTables(normalTableList, errorMsg_Event);
                BuildNormalTableCellNewValue(normalTableList, errorMsg_Event);
                foreach (var normalTable in normalTableList)
                {
                    normalTable.TableNumber = tableList.Count + 1;
                    tableList.Add(normalTable);
                }
                tableList = tableList.OrderBy(o => o.TableContentStartParagraphNumber).ToList();
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
        private static WordTable GetWordTable_old(Table wordTable)
        {

            var table = new WordTable();
            string tableXml = wordTable.Range.WordOpenXML;
            XDocument document = XDocument.Parse(tableXml);
            var rows = document.Root.Descendants().Where(f => f.Name.LocalName == "tr").ToList();
            //所有行都是空行的表格 返回空
            if (rows.All(w => string.IsNullOrWhiteSpace(w.Value)))
            {
                return null;
            }
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
                        StartRowIndex = table.Rows.Count + 1,
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
                ////如果整行内容都是空 不计入行
                //if (rowCells.Select(s => s.OldValue.Trim()).All(w => string.IsNullOrWhiteSpace(w)))
                //{
                //    continue;
                //}
                var tableRow = new WordTableRow()
                {
                    RowNumber = table.Rows.Count + 1,
                    RowCells = rowCells,
                };
                try
                {
                    Row wordRow = wordTable.Rows[table.Rows.Count + 1];
                    tableRow.Range = wordRow.Range;
                }
                catch { }
                table.Rows.Add(tableRow);
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

                if (row.Range == null)
                {
                    var rowCellList = row.RowCells.Where(w => w.Range != null).OrderBy(o => o.StartColumnIndex).ToList();
                    var firstCellRange = rowCellList.FirstOrDefault().Range.Duplicate;
                    firstCellRange.End = rowCellList.LastOrDefault().Range.End;
                    string text = firstCellRange.Text;
                    row.Range = firstCellRange;

                }
            }
            //第一个单元格垂直合并数量
            int firstCellRowSpan = table.Rows.FirstOrDefault().RowCells.FirstOrDefault().RowSpan;
            //lxz 判断是否有【人民币元】和空行，
            foreach (var row in table.Rows)
            {
                #region lxz 添加判断 人名币 和 空行 认为是表头
                ////判断当前表格中是否有人民币行，如果并且大于firstCellRowSpan则firstCellRowSpan设置为当前行数
                //var rmbCount = row.RowCells.Where(x => x.OldValue.Equals("人民币") || x.OldValue.Equals("人民币元")).Count();
                //if (rmbCount > 1)
                //{
                //    if (row.RowNumber > firstCellRowSpan)
                //    {
                //        firstCellRowSpan = row.RowNumber;
                //    }
                //    //判断人民币下一行是否空行
                //    if (row.RowNumber + 1 < table.Rows.Count)
                //    {
                //        var isNotEmpty = table.Rows[row.RowNumber + 1].RowCells.Where(x => !string.IsNullOrEmpty(x.OldValue)).Any();
                //        if (!isNotEmpty)
                //        {
                //            firstCellRowSpan = row.RowNumber + 1;
                //        }
                //    }
                //}
                #endregion

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
            //lxz 添加判断 人名币 和 空行 认为是表头
            SupplementRMBHeader(table);
            return table;

        }

        private static WordTable GetWordTable(Table wordTable)
        {
            int wdActiveEndPageNumber = -1;
            try
            {
                var _paragraph = wordTable.Range.Paragraphs.First;
                wdActiveEndPageNumber = Convert.ToInt32(_paragraph.Range.Information[WdInformation.wdActiveEndPageNumber]);

                var table = new WordTable();
                string tableXml = wordTable.Range.WordOpenXML;
                XDocument document = XDocument.Parse(tableXml);
                var rows = document.Root.Descendants().Where(f => f.Name.LocalName == "tr").ToList();
                //所有行都是空行的表格 返回空
                if (rows.All(w => string.IsNullOrWhiteSpace(w.Value)))
                {
                    return null;
                }

                var temp_rowList = new List<WordTableRow>();
                var rdx = 1;
                foreach (var row in rows)
                {
                    var rowCells = new List<WordTableCell>();
                    temp_rowList.Add(new WordTableRow
                    {
                        RowNumber = rdx,
                        RowCells = rowCells,
                    });
                    var x_cells = row.Descendants().Where(w => w.Name.LocalName == "tc").ToList();
                    int columnIndex = 1;
                    foreach (var x_cell in x_cells)
                    {
                        var tcPr = x_cell.Descendants().FirstOrDefault(w => w.Name.LocalName == "tcPr");
                        var gridSpan = tcPr.Descendants().FirstOrDefault(w => w.Name.LocalName == "gridSpan");
                        var gridSpanVal = gridSpan == null ? 1 : Convert.ToInt32(gridSpan.Attributes().FirstOrDefault(w => w.Name.LocalName == "val").Value);

                        var vMerge = tcPr.Descendants().FirstOrDefault(w => w.Name.LocalName == "vMerge");

                        //var vMergeVal = vMerge == null ? null : vMerge.Attributes().FirstOrDefault(w => w.Name.LocalName == "val")?.Value ?? string.Empty;

                        //1:Restart 2:Continue 3:NoMerge
                        var vMergeVal = "NoMerge";
                        if (vMerge != null)
                        {
                            vMergeVal = vMerge.Attributes().FirstOrDefault(w => w.Name.LocalName == "val")?.Value ?? "Continue";
                        }
                        var tableCell = new WordTableCell()
                        {
                            OldValue = x_cell.Value,
                            StartRowIndex = rdx,
                            ActualStartColumnIndex = x_cells.IndexOf(x_cell) + 1,
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
                    rdx++;

                }

                //计算单元格合并行
                var tblColCount = temp_rowList.Max(x => x.RowCells.Count);
                var tblRowCount = temp_rowList.Count;
                for (var colIdx = 0; colIdx < tblColCount; colIdx++)
                {
                    var colList = new List<WordTableCell>();
                    for (int _rdx = 0; _rdx < tblRowCount; _rdx++)
                    {
                        var _cell = temp_rowList[_rdx].RowCells.Where(x => x.StartColumnIndex == colIdx + 1).FirstOrDefault();
                        if (_cell != null)
                        {
                            colList.Add(_cell);
                        }
                    }

                    //var fristRelativeRowIndex = -1;
                    //var lastRelativeRowIndex = -1;
                    //var rowSpan = -1;
                    //进行行合并
                    for (var rowIdx = 0; rowIdx < tblRowCount; rowIdx++)
                    {
                        var cell = temp_rowList[rowIdx].RowCells.Where(x => x.StartColumnIndex == colIdx + 1).FirstOrDefault();
                        if (cell == null)
                        {
                            continue;
                        }
                        cell.RowSpan = 1;
                        //1:Restart 2:Continue 3:NoMerge
                        if (cell.VMergeVal.ToUpper() == "Restart".ToUpper())
                        {
                            var lastCell = colList.Where(x => x.StartRowIndex > cell.StartRowIndex && x.VMergeVal.ToUpper() != "Continue".ToUpper()).FirstOrDefault();
                            //rowSpan = lastCell.StartRowIndex - cell.StartRowIndex;
                            //fristRelativeRowIndex = cell.RowIndex;
                            //lastRelativeRowIndex = lastCell.RowIndex - 1;

                            //cell.StartRowIndex =  fristRelativeRowIndex;
                            //cell.EndRowIndex = lastRelativeRowIndex;

                            cell.RowSpan = lastCell.StartRowIndex - cell.StartRowIndex;// rowSpan;
                        }
                        else if (cell.VMergeVal.ToUpper() == "Continue".ToUpper())
                        {
                            cell.RowSpan = 1;//rowSpan
                        }
                        else if (cell.VMergeVal.ToUpper() == "NoMerge".ToUpper())
                        {
                            cell.RowSpan = 1;
                            //rowSpan = -1;
                        }
                    }
                }

                var sindex = 0;
                var removeRowIndexList = new List<int>();
                for (int rowIndex = 0; rowIndex < tblRowCount; rowIndex++)
                {
                    var row = temp_rowList[rowIndex];
                    var isRemoveRow = false;
                    if (row.RowCells.Where(x => x.VMergeVal.ToUpper() == "Continue".ToUpper()).Count() == row.RowCells.Count())
                    {
                        //删除该行
                        removeRowIndexList.Add(rowIndex);
                        isRemoveRow = true;
                        sindex++;
                    }
                    if (!isRemoveRow)
                    {
                        row.RowCells.ForEach(c =>
                        {
                            c.StartRowIndex -= sindex;
                            if (!isRemoveRow)
                            {
                                if (c.VMergeVal.ToUpper() != "Continue".ToUpper())
                                {
                                    c.Range = wordTable.Cell(c.StartRowIndex, c.ActualStartColumnIndex)?.Range;
                                }
                            }
                        });
                        row.RowNumber -= sindex;
                        if (!isRemoveRow && !row.RowCells.Where(x => x.ColSpan > 1 || x.RowSpan > 1).Any())
                        {
                            var rowCellList = row.RowCells.Where(w => w.Range != null).OrderBy(o => o.StartColumnIndex).ToList();
                            var firstCellRange = rowCellList.FirstOrDefault().Range.Duplicate;
                            firstCellRange.End = rowCellList.LastOrDefault().Range.End;
                            row.Range = firstCellRange;
                            row.IsMatchRowRange = row.Range != null;
                        }
                    }
                }

                if (removeRowIndexList.Any())
                {
                    var count = removeRowIndexList.Count - 1;
                    for (int i = count; i >= 0; i--)
                    {
                        var idx = removeRowIndexList[i];
                        temp_rowList.RemoveAt(idx);
                    }
                }
                table.Rows = temp_rowList;
                //第一个单元格垂直合并数量
                int firstCellRowSpan = table.Rows.FirstOrDefault().RowCells.FirstOrDefault().RowSpan;
                //lxz 判断是否有【人民币元】和空行，
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
                    row.RowCells = row.RowCells.Where(w => w.VMergeVal.ToUpper() != "Continue".ToUpper()).ToList();
                }
                //lxz 添加判断 人名币 和 空行 认为是表头
                SupplementRMBHeader(table);

                return table;
            }
            catch (Exception ex)
            {
                var msg = "解析表格失败";
                if (wdActiveEndPageNumber > 0)
                {
                    msg += $",第{wdActiveEndPageNumber}页";
                }
                throw new Exception($"{msg};ex:{ex.Message}", ex);
            }
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

        /// <summary>
        /// 生成正常表格新值
        /// </summary>
        /// <param name="tables"></param>
        private static void BuildNormalTableCellNewValue(List<WordTable> tables, Action<object, Events.NodeNotifyEventArgs> errorMsg_Event = null)
        {
            var replaceItemList = WordTableConfigHelper.GetCellReplaceItemConfig();
            int lastTableIndex = tables.IndexOf(tables.LastOrDefault());
            for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
            {
                string errorMsg = string.Empty;
                var table = tables[tableIndex];
                try
                {

                    #region 同表左右替换 判断当前表格所有表头是否包含两个及以上不同日期或者包含任意一组关键字

                    var horizontalHeadRowCellList = GetHorizontalMergeTableHeadRow(table.HeadRows);

                    var horizontalDateReplaceMatchItemList = horizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                    && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date).ToList();
                    var horizontalDateReplaceMatchItemGroupCount = horizontalDateReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                    var horizontalKeywordReplaceMatchItemList = horizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                    && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Keyword).ToList();
                    var filterHorizontalKeywordReplaceMatchItemList = new List<ReplaceCell>();
                    horizontalKeywordReplaceMatchItemList.ForEach(matchItem =>
                    {
                        var matchItemKeyvaluePair = replaceItemList.FirstOrDefault(w => w.Key == matchItem.ReplaceMatchItem || w.Value == matchItem.ReplaceMatchItem);
                        bool isIncludeKeyvaluePair = new string[] { matchItemKeyvaluePair.Key, matchItemKeyvaluePair.Value }.All(w => horizontalKeywordReplaceMatchItemList.Select(s => s.ReplaceMatchItem).Contains(w));
                        if (isIncludeKeyvaluePair)
                        {
                            filterHorizontalKeywordReplaceMatchItemList.Add(matchItem);
                        }
                    });
                    horizontalKeywordReplaceMatchItemList = filterHorizontalKeywordReplaceMatchItemList;
                    var horizontalKeywordReplaceMatchItemGroupCount = horizontalKeywordReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                    if (horizontalDateReplaceMatchItemGroupCount >= 2 ||
                       horizontalKeywordReplaceMatchItemGroupCount >= 2)
                    {
                        //执行同表跨列替换逻辑
                        SameTableCrossColumnReplace(table, horizontalDateReplaceMatchItemList, horizontalKeywordReplaceMatchItemList);
                        table.OperationType = OperationTypeEnum.ReplaceText;

                        //lxz 2024-07-24 添加逻辑
                        ReSetTableOperationType(table, horizontalHeadRowCellList, null);
                        continue;
                    }
                    #endregion

                    #region 同表上下替换 判断当前表格第一列是否包含两个及以上不同日期或者包含任意一组关键字
                    var verticalHeadRowCellList = GetVerticalTableHeadRow(table.Rows);
                    var isTheSameHeadRow = IsTheSameHeadRow(table, tables);
                    if (!isTheSameHeadRow)
                    {
                        var verticalDateReplaceMatchItemList = verticalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                         && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date).ToList();
                        var verticalDateReplaceMatchItemGroupCount = verticalDateReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                        var verticalKeywordReplaceMatchItemList = verticalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                        && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Keyword).ToList();
                        var filterVerticalKeywordReplaceMatchItemList = new List<ReplaceCell>();
                        verticalKeywordReplaceMatchItemList.ForEach(matchItem =>
                        {
                            var matchItemKeyvaluePair = replaceItemList.FirstOrDefault(w => w.Key == matchItem.ReplaceMatchItem || w.Value == matchItem.ReplaceMatchItem);
                            bool isIncludeKeyvaluePair = new string[] { matchItemKeyvaluePair.Key, matchItemKeyvaluePair.Value }.All(w => verticalKeywordReplaceMatchItemList.Select(s => s.ReplaceMatchItem).Contains(w));
                            if (isIncludeKeyvaluePair)
                            {
                                filterVerticalKeywordReplaceMatchItemList.Add(matchItem);
                            }
                        });
                        verticalKeywordReplaceMatchItemList = filterVerticalKeywordReplaceMatchItemList;
                        var verticalKeywordReplaceMatchItemGroupCount = verticalKeywordReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                        if (verticalDateReplaceMatchItemGroupCount >= 2 ||
                           verticalKeywordReplaceMatchItemGroupCount >= 2)
                        {
                            //执行同表跨行替换逻辑
                            SameTableCrossRowReplace(table, verticalDateReplaceMatchItemList, verticalKeywordReplaceMatchItemList);
                            table.OperationType = OperationTypeEnum.ReplaceText;

                            //lxz 2024-07-24 添加逻辑
                            ReSetTableOperationType(table, horizontalHeadRowCellList, verticalHeadRowCellList);
                            continue;
                        }
                    }

                    #endregion

                    #region 跨表上下替换 只有当前表非最后一个表格且与下一个表格表头除了匹配项完全一致 
                    if (tableIndex < lastTableIndex)
                    {

                        var nextTable = tables[tableIndex + 1];

                        //两个表格所在页码数最多差一页
                        if (Math.Abs(table.PageNumber - nextTable.PageNumber) <= 1)
                        {
                            //判断当前表格与下一个表格是否表头除开日期部分是否完全一致 上下两个表格均有一个匹配项
                            //且当前表格第一列所有行中内容是否存在任意一项在下一个表格第一列所有行中存在
                            var nextTableHorizontalHeadRowCellList = nextTable != null ? GetHorizontalMergeTableHeadRow(nextTable.HeadRows) : new List<ReplaceCell>();

                            var nextTableHorizontalDateReplaceMatchItemList = nextTableHorizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                             && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date).ToList();
                            var nextTableHorizontalDateReplaceMatchItemGroupCount = nextTableHorizontalDateReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                            var nextTableHorizontalKeywordReplaceMatchItemList = nextTableHorizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                              && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Keyword).ToList();
                            var nextTableHorizontalKeywordReplaceMatchItemGroupCount = nextTableHorizontalKeywordReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                            string tableHeadRowContent = string.Join("", horizontalHeadRowCellList.Select(s => s.CellValue)).ReplaceDate();
                            string nextTableHeadRowContent = string.Join("", nextTableHorizontalHeadRowCellList.Select(s => s.CellValue)).ReplaceDate();

                            string currentTableFirstReplaceMatchItem = horizontalKeywordReplaceMatchItemList.FirstOrDefault()?.ReplaceMatchItem;
                            string nextTableFirstReplaceMatchItem = nextTableHorizontalKeywordReplaceMatchItemList.FirstOrDefault()?.ReplaceMatchItem;
                            bool isKeyValuePair = !string.IsNullOrWhiteSpace(currentTableFirstReplaceMatchItem) && !string.IsNullOrWhiteSpace(nextTableFirstReplaceMatchItem)
                                && replaceItemList.Count(w => (w.Key == currentTableFirstReplaceMatchItem && w.Value == nextTableFirstReplaceMatchItem) ||
                                (w.Key == nextTableFirstReplaceMatchItem && w.Value == currentTableFirstReplaceMatchItem)) == 1;

                            if (tableHeadRowContent == nextTableHeadRowContent //上下两个表除开日期部分表头一致
                                && ((horizontalDateReplaceMatchItemGroupCount == 1 && nextTableHorizontalDateReplaceMatchItemGroupCount == 1)//上下两个表都有一个日期匹配项
                                || (horizontalKeywordReplaceMatchItemGroupCount == 1 && nextTableHorizontalKeywordReplaceMatchItemGroupCount == 1 && isKeyValuePair))) //上下两个表都有一个关键字匹配项，且是一堆键值对
                            {
                                //执行跨表替换逻辑
                                CrossTableReplace(table, horizontalDateReplaceMatchItemList, horizontalKeywordReplaceMatchItemList,
                                    nextTable, nextTableHorizontalDateReplaceMatchItemList, nextTableHorizontalKeywordReplaceMatchItemList);
                                table.OperationType = OperationTypeEnum.ReplaceText;
                                nextTable.OperationType = OperationTypeEnum.ReplaceText;
                                //循环跳过下一个表
                                tableIndex++;

                            }
                        }

                    }

                    #endregion

                    //lxz 2024-07-01 添加逻辑
                    ReSetTableOperationType(table, horizontalHeadRowCellList, verticalHeadRowCellList);

                    ////lxz 2024-07-01 添加逻辑
                    ////表格表头包含年份，却没有执行上面的替换逻辑，则表头应该替换颜色
                    //if (table.OperationType == OperationTypeEnum.NotOperation && (
                    //      horizontalHeadRowCellList.Any(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem))
                    //      || verticalHeadRowCellList.Any(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem))
                    //      || horizontalHeadRowCellList.Any(w => w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Disturb)
                    //      ))
                    //{
                    //    table.OperationType = OperationTypeEnum.ChangeColor;
                    //}
                    //else if (table.OperationType == OperationTypeEnum.ReplaceText && !table.Rows.Where(x => x.RowCells.Any(c => c.IsReplaceValue)).Any())
                    //{
                    //    table.OperationType = OperationTypeEnum.ChangeColor;
                    //}
                }
                catch (Exception ex)
                {
                    errorMsg = $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})生成单元格新值失败，{ex.Message}";
                    table.OperationType = OperationTypeEnum.ChangeColor;
                    table.ErrorMsgs.Add(errorMsg);
                    errorMsg.Console(ConsoleColor.Red);
                    ex.StackTrace.Console(ConsoleColor.Red);
                    if (errorMsg_Event != null)
                    {
                        errorMsg_Event(null, new Events.NodeNotifyEventArgs
                        {
                            Type = 2,
                            Message = errorMsg
                        });
                    }
                }
            }
        }

        #endregion

        #region 制表位表格


        /// <summary>
        /// 获取识别失败制表位表格
        /// </summary>
        /// <param name="notUseParagraphs"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        private static List<WordTable> GetIdentifyFailTabStopTableList(List<WordParagraph> notUseParagraphs, List<WordTable> tables)
        {
            var identifyFailTabStopTableList = new List<WordTable>();

            var tabStopTableParagraphList = new List<List<WordParagraph>>();

            string strRule1 = @"^[一-十|四]{1,3}、\t";
            string strRule2 = @"^\d+(\.+){1,}";//以数字+.开头
            string strRule3 = @"^\([a-z0-9]+\)\t";//以数字+.开头

            var paragraphList = new List<WordParagraph>();
            var prevParagraph = notUseParagraphs.FirstOrDefault();
            paragraphList.Add(prevParagraph);
            var isTitle = false;
            for (int i = 1; i < notUseParagraphs.Count; i++)
            {
                var currentParagraph = notUseParagraphs[i];

                //lxz 2024-07-25 添加判断标题逻辑
                var paraOldText = currentParagraph.OldText;
                if (Regex.IsMatch(paraOldText, strRule1)
                    || Regex.IsMatch(paraOldText, strRule2)
                    || Regex.IsMatch(paraOldText, strRule3))
                {
                    isTitle = true;
                    prevParagraph = currentParagraph;
                    continue;
                }
                if (isTitle)
                {
                    isTitle = false;
                    tabStopTableParagraphList.Add(paragraphList);

                    paragraphList = new List<WordParagraph>() { currentParagraph };
                }
                else if (prevParagraph.ParagraphNumber + 1 == currentParagraph.ParagraphNumber)
                {
                    //连续段落
                    paragraphList.Add(currentParagraph);
                }
                else
                {
                    //非连续段落
                    if (paragraphList.Count >= 2)
                    {
                        tabStopTableParagraphList.Add(paragraphList);
                    }
                    paragraphList = new List<WordParagraph>() { currentParagraph };

                }
                prevParagraph = currentParagraph;
            }

            //排除连续段落 包含\t段落数量少于2个段落的
            tabStopTableParagraphList = tabStopTableParagraphList.Where(w => w.Count(ww => ww.OldText.Contains("\t")) >= 2).ToList();

            foreach (var tableParagraphList in tabStopTableParagraphList)
            {

                var tableList = FindTables(tableParagraphList);
                if (tableList.Any())
                {
                    identifyFailTabStopTableList.AddRange(tableList);
                }
            }

            foreach (var table in identifyFailTabStopTableList)
            {
                MatchTabStopTableCellRange(table);

                int maxRowCellCount = table.Rows.Max(m => m.RowCells.Count);
                if (table.HeadRows.Any(w => w.RowCells.Count != maxRowCellCount))
                {
                    //如果表头行存在列数与最大列数不一致 代表有合并表头 不对有合并表头的计算制表位表格进行替换
                    table.OperationType = OperationTypeEnum.ChangeColor;
                    continue;
                }
            }

            return identifyFailTabStopTableList;
        }


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
                            OldValue = cell.OldValue.Trim(),
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

            //单元格包含多个匹配项表示为干扰项
            if (cellValue.GetAllReplaceItemList().Count > 1)
            {
                replaceMatchItem = null;
                replaceMatchItemType = ReplaceMatchItemTypeEnum.Disturb;
            }
            return (replaceMatchItem, replaceMatchItemType);
        }

        /// <summary>
        /// 同表跨列替换
        /// </summary>
        /// <param name="table"></param>
        /// <param name="dateReplaceMatchItems"></param>
        /// <param name="keywordReplaceMatchItems"></param>
        private static void SameTableCrossColumnReplace(WordTable table, List<ReplaceCell> dateReplaceMatchItems, List<ReplaceCell> keywordReplaceMatchItems)
        {
            var allCellList = table.Rows.SelectMany(s => s.RowCells).ToList();
            #region 日期
            if (dateReplaceMatchItems.Any() && dateReplaceMatchItems.Count >= 2)
            {
                var dateReplaceMatchItemGroupList = dateReplaceMatchItems.GroupBy(g => g.ReplaceMatchItemDate).ToList();
                if (dateReplaceMatchItems.Count % 2 == 0 && dateReplaceMatchItemGroupList.All(w => w.Count() >= 2))
                {
                    //日期两两一组重复出现 如：2023年 2022年 2023年 2022年
                    //匹配项数量是偶数 且匹配项存在重复 按照最近的两个匹配项为一组 
                    for (int i = 0; i < dateReplaceMatchItems.Count; i++)
                    {
                        var currentReplaceCell = dateReplaceMatchItems[i];
                        var nextReplaceCell = dateReplaceMatchItems[i + 1];

                        var currentMatchItemColumnCellList = allCellList.Where(w => w.StartColumnIndex == currentReplaceCell.Index).ToList();
                        var nextMatchItemColumnCellList = allCellList.Where(w => w.StartColumnIndex == nextReplaceCell.Index).ToList();

                        //判断替换方向 
                        if (currentReplaceCell.ReplaceMatchItemDate > nextReplaceCell.ReplaceMatchItemDate)
                        {

                            //当前日期匹配项大于下一个日期匹配项 从上往下替换
                            foreach (var cell in currentMatchItemColumnCellList)
                            {
                                if (cell.IsHeadColumn)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = "";
                                }
                                cell.IsReplaceValue = true;
                            }
                            foreach (var cell in nextMatchItemColumnCellList)
                            {
                                if (cell.IsHeadColumn)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = currentMatchItemColumnCellList.FirstOrDefault(w => w.StartRowIndex == cell.StartRowIndex)?.OldValue;
                                }
                                cell.IsReplaceValue = true;
                            }

                        }
                        else
                        {
                            //当前日期匹配项小于等于下一个日期匹配项 从下往上替换
                            foreach (var cell in nextMatchItemColumnCellList)
                            {
                                if (cell.IsHeadColumn)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = "";
                                }
                                cell.IsReplaceValue = true;
                            }
                            foreach (var cell in currentMatchItemColumnCellList)
                            {
                                if (cell.IsHeadColumn)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = nextMatchItemColumnCellList.FirstOrDefault(w => w.StartRowIndex == cell.StartRowIndex)?.OldValue;
                                }
                                cell.IsReplaceValue = true;
                            }

                        }
                        //下一行已经替换 跳过循环
                        i++;

                    }
                }
                else
                {
                    //根据表头新单元格值分组
                    var headCellValueGroupbyResultList = dateReplaceMatchItems.GroupBy(g => g.CellValue.Replace(g.ReplaceMatchItem, "")).ToList();
                    foreach (var headCellValueGroupbyResult in headCellValueGroupbyResultList)
                    {
                        if (headCellValueGroupbyResult.Count() <= 1)
                        {
                            continue;
                        }
                        //所有替换数据匹配列根据匹配日期降序排序 
                        var replaceCellGroupResultList = headCellValueGroupbyResult.OrderByDescending(o => o.ReplaceMatchItemDate).ToList();
                        //第一个匹配列数据清空
                        var firstMatchItemColumnCellList = allCellList.Where(w => w.StartColumnIndex == replaceCellGroupResultList.FirstOrDefault().Index).ToList();
                        foreach (var cell in firstMatchItemColumnCellList)
                        {
                            if (cell.IsHeadColumn)
                            {
                                string cellDateString = cell.OldValue.GetDateString();
                                if (!string.IsNullOrWhiteSpace(cellDateString))
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(replaceCellGroupResultList, cell.OldValue);
                                    cell.IsReplaceValue = true;
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrWhiteSpace(cell.OldValue))
                                {
                                    cell.NewValue = "";
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                        //其他列依次从左取前一列
                        for (int i = 1; i < replaceCellGroupResultList.Count; i++)
                        {
                            var currentReplaceHeadCell = replaceCellGroupResultList[i];
                            var prevReplaceHeadCell = replaceCellGroupResultList[i - 1];
                            //前一匹配列所有单元格
                            var prevMatchItemColumnCellList = allCellList.Where(w => w.StartColumnIndex == prevReplaceHeadCell.Index).ToList();
                            //当前匹配列所有单元格
                            var currentMatchItemColumnCellList = allCellList.Where(w => w.StartColumnIndex == currentReplaceHeadCell.Index).ToList();
                            foreach (var cell in currentMatchItemColumnCellList)
                            {
                                if (cell.IsHeadColumn)
                                {
                                    string cellDateString = cell.OldValue.GetDateString();
                                    if (!string.IsNullOrWhiteSpace(cellDateString))
                                    {
                                        cell.NewValue = GetNextMaxDateHeadCellValue(replaceCellGroupResultList, cell.OldValue);
                                        cell.IsReplaceValue = true;
                                    }
                                }
                                else
                                {
                                    var newValue = prevMatchItemColumnCellList.FirstOrDefault(w => w.StartRowIndex == cell.StartRowIndex)?.OldValue;
                                    if (cell.OldValue != newValue)
                                    {
                                        cell.NewValue = newValue;
                                        cell.IsReplaceValue = true;
                                    }
                                }

                            };
                        }

                    }
                }

            }
            #endregion

            #region 关键字
            if (keywordReplaceMatchItems.Any() && keywordReplaceMatchItems.Count >= 2)
            {
                var replaceItemList = WordTableConfigHelper.GetCellReplaceItemConfig();
                var keywordReplaceMatchItemGroupList = keywordReplaceMatchItems.GroupBy(g => g.ReplaceMatchItem).ToList();
                var replaceMatchItemList = keywordReplaceMatchItems.Select(s => s.ReplaceMatchItem).ToList();
                var hasRepeatedPair = replaceMatchItemList.HasRepeatedPair();
                if (keywordReplaceMatchItems.Count % 2 == 0 && keywordReplaceMatchItemGroupList.All(w => w.Count() >= 2) &&
                    keywordReplaceMatchItemGroupList.Count == 2 && hasRepeatedPair)
                {
                    //关键字两两一组重复出现 如：年末数 年初数 年末数 年初数
                    //匹配项数量是偶数 且匹配项存在重复 按照最近的两个匹配项为一组 
                    for (int i = 0; i < keywordReplaceMatchItems.Count; i++)
                    {
                        var currentReplaceCell = keywordReplaceMatchItems[i];
                        var nextReplaceCell = keywordReplaceMatchItems[i + 1];

                        //验证相邻两个关键字是否是同一对键值对
                        var matchKeyValuePairCount = replaceItemList.Count(w => (w.Key == currentReplaceCell.ReplaceMatchItem && w.Value == nextReplaceCell.ReplaceMatchItem)
                        || (w.Key == nextReplaceCell.ReplaceMatchItem && w.Value == currentReplaceCell.ReplaceMatchItem));
                        if (matchKeyValuePairCount != 1)
                        {
                            //相邻两组关键字不是键值对 跳过当前行和下一行
                            i++;
                            continue;
                        }
                        var currentMatchItemColumnCellList = allCellList.Where(w => w.StartColumnIndex == currentReplaceCell.Index).ToList();
                        var nextMatchItemColumnCellList = allCellList.Where(w => w.StartColumnIndex == nextReplaceCell.Index).ToList();

                        if (replaceItemList.Any(w => w.Key == currentReplaceCell.ReplaceMatchItem))
                        {
                            //当前匹配列是key 从左往右替换
                            foreach (var cell in currentMatchItemColumnCellList)
                            {
                                if (!cell.IsHeadColumn)
                                {
                                    cell.NewValue = "";
                                    cell.IsReplaceValue = true;
                                }

                            }
                            foreach (var cell in nextMatchItemColumnCellList)
                            {
                                if (!cell.IsHeadColumn)
                                {
                                    cell.NewValue = currentMatchItemColumnCellList.FirstOrDefault(w => w.StartRowIndex == cell.StartRowIndex)?.OldValue;
                                    cell.IsReplaceValue = true;
                                }
                            }

                        }
                        else
                        {
                            //下一匹配列是key 从右往左替换
                            foreach (var cell in nextMatchItemColumnCellList)
                            {
                                if (!cell.IsHeadColumn)
                                {
                                    cell.NewValue = "";
                                    cell.IsReplaceValue = true;
                                }
                            }
                            foreach (var cell in currentMatchItemColumnCellList)
                            {
                                if (!cell.IsHeadColumn)
                                {
                                    cell.NewValue = nextMatchItemColumnCellList.FirstOrDefault(w => w.StartRowIndex == cell.StartRowIndex)?.OldValue;
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                        i++;
                    }
                }
                else
                {
                    var alreadyReplaceMatchItemList = new List<string>();
                    foreach (var replaceCell in keywordReplaceMatchItems)
                    {
                        var replaceItem = replaceItemList.FirstOrDefault(w => w.Key == replaceCell.ReplaceMatchItem || w.Value == replaceCell.ReplaceMatchItem);

                        /*
                         * lxz 2024-07-03 修改逻辑
                         * 根据当前单元个判断是key 还是 val 
                         * 在根据当前单元格名称，替换key val值，的得到对应的单元格，排除使用过的列号
                         * 判断keycell和valcell 是否获取到，
                         * 记录使用过的列号
                         */

                        ReplaceCell keyReplaceCell = null;
                        ReplaceCell valueReplaceCell = null;
                        var alreadylist = alreadyReplaceMatchItemList.SelectMany(x => x.Split('_')).ToList();
                        if (replaceCell.ReplaceMatchItem == replaceItem.Key)
                        {
                            keyReplaceCell = replaceCell;
                            //var val_Item = keyReplaceCell.CellValue.Replace(replaceItem.Key, replaceItem.Value).Trim();
                            //年末数金额 年初数金额(已重述)
                            valueReplaceCell = keywordReplaceMatchItems.Where(x => !alreadylist.Contains(x.Index.ToString()))
                                .FirstOrDefault(w => w.ReplaceMatchItem == replaceItem.Value && w.CellValue.ReplaceEmpty(w.ReplaceMatchItem).IsContains(keyReplaceCell.CellValue.ReplaceEmpty(keyReplaceCell.ReplaceMatchItem)));
                        }
                        else if (replaceCell.ReplaceMatchItem == replaceItem.Value)
                        {
                            valueReplaceCell = replaceCell;
                            //var key_Item = valueReplaceCell.CellValue.Replace(replaceItem.Value, replaceItem.Key).Trim();
                            keyReplaceCell = keywordReplaceMatchItems.Where(x => !alreadylist.Contains(x.Index.ToString()))
                                .FirstOrDefault(w => w.ReplaceMatchItem == replaceItem.Key && w.CellValue.ReplaceEmpty(w.ReplaceMatchItem).IsContains(valueReplaceCell.CellValue.ReplaceEmpty(valueReplaceCell.ReplaceMatchItem)));
                        }

                        if (keyReplaceCell != null && valueReplaceCell != null)
                        {
                            //lxz 2024-07-03 判断【key的列号_val的列号】 
                            //if (!alreadyReplaceMatchItemList.Contains(keyReplaceCell.ReplaceMatchItem + "_" + valueReplaceCell.ReplaceMatchItem))

                            if (!alreadyReplaceMatchItemList.Contains(keyReplaceCell.Index + "_" + valueReplaceCell.Index))
                            {
                                var matchKeyColumnCellList = allCellList.Where(w => w.StartColumnIndex == keyReplaceCell.Index && !w.IsHeadColumn).ToList();
                                var matchValueColumnCellList = allCellList.Where(w => w.StartColumnIndex == valueReplaceCell.Index && !w.IsHeadColumn).ToList();
                                foreach (var cell in matchKeyColumnCellList)
                                {
                                    //lxz 临时处理，由于表格识别错误，tableNumber=55 多个表格识别成一个表格
                                    if (Regex.IsMatch(cell.OldValue, @"上年|本年|上期|本期|期初|期末|人民币"))
                                    {
                                        continue;
                                    }
                                    if (cell.OldValue == "")
                                    {
                                        continue;
                                    }
                                    cell.NewValue = "";
                                    cell.IsReplaceValue = true;
                                }
                                foreach (var cell in matchValueColumnCellList)
                                {
                                    //lxz 临时处理，由于表格识别错误，tableNumber=55 多个表格识别成一个表格
                                    if (Regex.IsMatch(cell.OldValue, @"上年|本年|上期|本期|期初|期末|人民币"))
                                    {
                                        continue;
                                    }

                                    var dataSourceCell = matchKeyColumnCellList.FirstOrDefault(w => w.StartRowIndex == cell.StartRowIndex);
                                    if (cell.NewValue == dataSourceCell?.OldValue)
                                    {
                                        continue;
                                    }
                                    cell.NewValue = dataSourceCell?.OldValue;
                                    cell.IsReplaceValue = true;
                                }
                                alreadyReplaceMatchItemList.Add($"{keyReplaceCell.Index}_{valueReplaceCell.Index}");
                            }

                        }
                    }
                }

            }
            #endregion
        }

        /// <summary>
        /// 同表跨行替换
        /// </summary>
        /// <param name="table"></param>
        /// <param name="dateReplaceMatchItems"></param>
        /// <param name="keywordReplaceMatchItems"></param>
        private static void SameTableCrossRowReplace(WordTable table, List<ReplaceCell> dateReplaceMatchItems, List<ReplaceCell> keywordReplaceMatchItems)
        {
            var allCellList = table.Rows.SelectMany(s => s.RowCells).ToList();
            #region 日期
            if (dateReplaceMatchItems.Any() && dateReplaceMatchItems.Count >= 2)
            {
                var dateReplaceMatchItemGroupList = dateReplaceMatchItems.GroupBy(g => g.ReplaceMatchItemDate).ToList();
                if (dateReplaceMatchItems.Count % 2 == 0 && dateReplaceMatchItemGroupList.All(w => w.Count() >= 2))
                {
                    //日期两两一组重复出现 如：2023年 2022年 2023年 2022年
                    //匹配项数量是偶数 且匹配项存在重复 按照最近的两个匹配项为一组 
                    int lastIndex = dateReplaceMatchItems.IndexOf(dateReplaceMatchItems.LastOrDefault());
                    for (int i = 0; i < dateReplaceMatchItems.Count; i++)
                    {
                        var currentReplaceCell = dateReplaceMatchItems[i];
                        var nextReplaceCell = dateReplaceMatchItems[i + 1];

                        var currentMatchItemRowCellList = allCellList.Where(w => w.StartRowIndex == currentReplaceCell.Index).ToList();
                        var nextMatchItemRowCellList = allCellList.Where(w => w.StartRowIndex == nextReplaceCell.Index).ToList();

                        //判断替换方向 
                        if (currentReplaceCell.ReplaceMatchItemDate < nextReplaceCell.ReplaceMatchItemDate)
                        {
                            //下一个匹配行日期大于当前日期 从下往上移动
                            foreach (var cell in nextMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex <= 1)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = "";
                                }
                                cell.IsReplaceValue = true;
                            }
                            foreach (var cell in currentMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex <= 1)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = nextMatchItemRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                }
                                cell.IsReplaceValue = true;
                            }

                        }
                        else
                        {
                            //下一个匹配行日期小于当前日期 从上往下移动
                            foreach (var cell in currentMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex <= 1)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = "";
                                }
                                cell.IsReplaceValue = true;
                            }
                            foreach (var cell in nextMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex <= 1)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(dateReplaceMatchItems, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = currentMatchItemRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                }
                                cell.IsReplaceValue = true;
                            }
                        }
                        //下一行已经替换 跳过循环
                        i++;

                    }
                }
                else
                {
                    //日期按顺序出现 如：2023年 2022年 2021年
                    //如果匹配到的是日期 要替换的表头移除掉匹配项文本 得到的表头单元格值应该是一致的
                    //dateReplaceMatchItems.ForEach(f => f.CellValue = f.CellValue.Replace(f.ReplaceMatchItem, ""));
                    //根据表头新单元格值分组
                    var headCellValueGroupbyResultList = dateReplaceMatchItems.GroupBy(g => g.CellValue.Replace(g.ReplaceMatchItem, "").RemoveWordTitle()).ToList();
                    foreach (var headCellValueGroupbyResult in headCellValueGroupbyResultList)
                    {
                        if (headCellValueGroupbyResult.Count() <= 1)
                        {
                            continue;
                        }
                        //所有替换数据匹配列根据匹配日期降序排序 
                        var replaceCellGroupResultList = headCellValueGroupbyResult.OrderByDescending(o => o.ReplaceMatchItemDate).ToList();
                        //更新当前分组第一个匹配行
                        var firstMatchItemRowCellList = allCellList.Where(w => w.StartRowIndex == replaceCellGroupResultList.FirstOrDefault().Index).ToList();
                        foreach (var cell in firstMatchItemRowCellList)
                        {
                            if (cell.StartColumnIndex <= 1)
                            {
                                cell.NewValue = GetNextMaxDateHeadCellValue(replaceCellGroupResultList, cell.OldValue);
                            }
                            else
                            {
                                cell.NewValue = "";
                            }
                            cell.IsReplaceValue = true;
                        };

                        //其他行 依次从上一个匹配项取值
                        for (int i = 1; i < replaceCellGroupResultList.Count; i++)
                        {
                            var currentReplaceHeadCell = replaceCellGroupResultList[i];
                            var prevReplaceHeadCell = replaceCellGroupResultList[i - 1];
                            //前一匹配行所有单元格
                            var prevMatchItemRowCellList = allCellList.Where(w => w.StartRowIndex == prevReplaceHeadCell.Index).ToList();
                            //当前匹配行所有单元格
                            var currentMatchItemRowCellList = allCellList.Where(w => w.StartRowIndex == currentReplaceHeadCell.Index).ToList();
                            foreach (var cell in currentMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex <= 1)
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(replaceCellGroupResultList, cell.OldValue);
                                }
                                else
                                {
                                    cell.NewValue = prevMatchItemRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;

                                }
                                cell.IsReplaceValue = true;
                            };
                        }

                    }
                }
            }
            #endregion

            #region 关键字
            if (keywordReplaceMatchItems.Any() && keywordReplaceMatchItems.Count >= 2)
            {
                var replaceItemList = WordTableConfigHelper.GetCellReplaceItemConfig();
                var keywordReplaceMatchItemGroupList = keywordReplaceMatchItems.GroupBy(g => g.ReplaceMatchItem).ToList();
                if (keywordReplaceMatchItems.Count % 2 == 0 && keywordReplaceMatchItemGroupList.All(w => w.Count() >= 2))
                {
                    //关键字两两一组重复出现 如：年末数 年初数 年末数 年初数
                    //匹配项数量是偶数 且匹配项存在重复 按照最近的两个匹配项为一组 
                    int lastIndex = keywordReplaceMatchItems.IndexOf(keywordReplaceMatchItems.LastOrDefault());
                    for (int i = 0; i < keywordReplaceMatchItems.Count; i++)
                    {
                        var currentReplaceCell = keywordReplaceMatchItems[i];
                        var nextReplaceCell = keywordReplaceMatchItems[i + 1];

                        //验证相邻两个关键字是否是同一对键值对
                        var matchKeyValuePairCount = replaceItemList.Count(w => (w.Key == currentReplaceCell.ReplaceMatchItem && w.Value == nextReplaceCell.ReplaceMatchItem)
                        || (w.Key == nextReplaceCell.ReplaceMatchItem && w.Value == currentReplaceCell.ReplaceMatchItem));
                        if (matchKeyValuePairCount != 1)
                        {
                            //相邻两组关键字不是键值对 跳过当前行和下一行
                            i++;
                            continue;
                        }
                        var currentMatchItemRowCellList = allCellList.Where(w => w.StartRowIndex == currentReplaceCell.Index).ToList();
                        var nextMatchItemRowCellList = allCellList.Where(w => w.StartRowIndex == nextReplaceCell.Index).ToList();

                        if (replaceItemList.Any(w => w.Key == currentReplaceCell.ReplaceMatchItem))
                        {
                            //当前匹配行是key 从上往下替换
                            //下一个匹配行日期小于当前日期 从上往下移动
                            foreach (var cell in currentMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex > 1)
                                {
                                    cell.NewValue = "";
                                    cell.IsReplaceValue = true;
                                }

                            }
                            foreach (var cell in nextMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex > 1)
                                {
                                    cell.NewValue = currentMatchItemRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                    cell.IsReplaceValue = true;
                                }
                            }

                        }
                        else
                        {
                            //下一匹配行是key 从下往上替换
                            //下一个匹配行日期大于当前日期 从下往上移动
                            foreach (var cell in nextMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex > 1)
                                {
                                    cell.NewValue = "";
                                    cell.IsReplaceValue = true;
                                }
                            }
                            foreach (var cell in currentMatchItemRowCellList)
                            {
                                if (cell.StartColumnIndex > 1)
                                {
                                    cell.NewValue = nextMatchItemRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                        i++;
                    }
                }
                else
                {
                    //关键字成对出现 如：年末数 年初数
                    var alreadyReplaceMatchItems = new List<string>();
                    foreach (var replaceCell in keywordReplaceMatchItems)
                    {
                        var replaceItem = replaceItemList.FirstOrDefault(w => w.Key == replaceCell.ReplaceMatchItem || w.Value == replaceCell.ReplaceMatchItem);
                        //匹配项key所在行第一个单元格
                        var keyReplaceCell = keywordReplaceMatchItems.FirstOrDefault(w => w.ReplaceMatchItem == replaceItem.Key);
                        //匹配项value所在行第一个单元格
                        var valueReplaceCell = keywordReplaceMatchItems.FirstOrDefault(w => w.ReplaceMatchItem == replaceItem.Value);
                        if (keyReplaceCell != null && valueReplaceCell != null)
                        {
                            if (!alreadyReplaceMatchItems.Contains(keyReplaceCell.ReplaceMatchItem + "_" + valueReplaceCell.ReplaceMatchItem))
                            {
                                var matchKeyRowCellList = allCellList.Where(w => w.StartRowIndex == keyReplaceCell.Index).ToList();
                                var matchValueRowCellList = allCellList.Where(w => w.StartRowIndex == valueReplaceCell.Index).ToList();
                                foreach (var cell in matchKeyRowCellList)
                                {
                                    if (cell.StartColumnIndex > 1)
                                    {
                                        cell.NewValue = "";
                                        cell.IsReplaceValue = true;
                                    }
                                }
                                foreach (var cell in matchValueRowCellList)
                                {
                                    if (cell.StartColumnIndex > 1)
                                    {
                                        cell.NewValue = matchKeyRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                        cell.IsReplaceValue = true;
                                    }
                                }
                                alreadyReplaceMatchItems.Add(keyReplaceCell.ReplaceMatchItem + "_" + valueReplaceCell.ReplaceMatchItem);
                            }

                        }
                    }
                }
            }
            #endregion

        }

        /// <summary>
        /// 跨表替换(只支持上下两个表各一个日期匹配项或者各一个键值对匹配项)
        /// </summary>
        /// <param name="table"></param>
        /// <param name="currentTableDateReplaceMatchItems"></param>
        /// <param name="currentTableKeywordReplaceMatchItems"></param>
        /// <param name="nextWordTable"></param>
        /// <param name="nextTableDateReplaceMatchItems"></param>
        /// <param name="nextTableKeywordReplaceMatchItems"></param>
        private static void CrossTableReplace(WordTable table, List<ReplaceCell> currentTableDateReplaceMatchItems,
            List<ReplaceCell> currentTableKeywordReplaceMatchItems, WordTable nextWordTable,
            List<ReplaceCell> nextTableDateReplaceMatchItems, List<ReplaceCell> nextTableKeywordReplaceMatchItems)
        {
            var tableDateRowCellList = table.DataRows.SelectMany(s => s.RowCells).ToList();
            var nextTableDateRowCellList = nextWordTable.DataRows.SelectMany(s => s.RowCells).ToList();

            #region 日期
            if (currentTableDateReplaceMatchItems.Any() && nextTableDateReplaceMatchItems.Any())
            {
                var currentTableMatchItemDate = currentTableDateReplaceMatchItems.FirstOrDefault().ReplaceMatchItemDate.Value;
                var nextTableMatchItemDate = nextTableDateReplaceMatchItems.FirstOrDefault().ReplaceMatchItemDate.Value;
                if (currentTableMatchItemDate > nextTableMatchItemDate)
                {
                    //当前表格匹配日期大于下一个表格匹配日期 从上往下替换
                    //清空当前表格 替换下一个表格
                    foreach (var row in table.Rows)
                    {
                        if (row.IsHeadRow)
                        {
                            //表头行 替换带日期的单元格值
                            foreach (var cell in row.RowCells)
                            {
                                string cellDateString = cell.OldValue.GetDateString();
                                if (!string.IsNullOrWhiteSpace(cellDateString))
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(new List<ReplaceCell> { currentTableDateReplaceMatchItems.FirstOrDefault(), nextTableDateReplaceMatchItems.FirstOrDefault() }, cell.OldValue);
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                        else
                        {
                            //数据行 清空数据
                            //判断当前表格数据行第一列内容是否存在于下一个表格数据行第一列单元格中 
                            var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                            if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue)
                                && nextTableDateRowCellList.Any(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue))
                            {
                                //清空从第二列开始数据
                                foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                                {
                                    if (!string.IsNullOrWhiteSpace(cell.OldValue))
                                    {
                                        cell.NewValue = "";
                                        cell.IsReplaceValue = true;
                                    }

                                }
                            }
                        }

                    }
                    foreach (var row in nextWordTable.Rows)
                    {
                        if (row.IsHeadRow)
                        {
                            //表头行 替换带日期的单元格值
                            foreach (var cell in row.RowCells)
                            {
                                string cellDateString = cell.OldValue.GetDateString();
                                if (!string.IsNullOrWhiteSpace(cellDateString))
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(new List<ReplaceCell> { currentTableDateReplaceMatchItems.FirstOrDefault(), nextTableDateReplaceMatchItems.FirstOrDefault() }, cell.OldValue);
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                        else
                        {
                            //数据行 替换数据
                            var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                            if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue))
                            {
                                var mapDataRowIndex = tableDateRowCellList.FirstOrDefault(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue)?.StartRowIndex ?? -1;
                                if (mapDataRowIndex > 1)
                                {
                                    //下一个表格当前数据行第一列内容在当前表格数据行第一列中存在
                                    //从第二列开始替换数据
                                    var dataRowCellList = tableDateRowCellList.Where(w => w.StartRowIndex == mapDataRowIndex).ToList();
                                    foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                                    {
                                        var newCellValue = dataRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                        if (cell.OldValue != newCellValue)
                                        {
                                            cell.NewValue = newCellValue;
                                            cell.IsReplaceValue = true;
                                        }

                                    }
                                }
                            }
                        }

                    }

                }
                else
                {
                    //当前表格匹配日期小于下一个表格匹配日期 从下往上替换
                    //清空下一个表格 替换当前表格
                    foreach (var row in nextWordTable.Rows)
                    {
                        if (row.IsHeadRow)
                        {
                            //表头行 替换带日期的单元格值
                            foreach (var cell in row.RowCells)
                            {
                                string cellDateString = cell.OldValue.GetDateString();
                                if (!string.IsNullOrWhiteSpace(cellDateString))
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(new List<ReplaceCell> { currentTableDateReplaceMatchItems.FirstOrDefault(), nextTableDateReplaceMatchItems.FirstOrDefault() }, cell.OldValue);
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                        else
                        {
                            //数据行 清空数据
                            var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                            if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue)
                                && tableDateRowCellList.Any(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue))
                            {
                                //清空从第二列开始数据
                                foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                                {
                                    if (!string.IsNullOrWhiteSpace(cell.OldValue))
                                    {
                                        cell.NewValue = "";
                                        cell.IsReplaceValue = true;
                                    }

                                }
                            }
                        }

                    }
                    foreach (var row in table.Rows)
                    {
                        if (row.IsHeadRow)
                        {
                            //表头行 替换带日期的单元格值
                            foreach (var cell in row.RowCells)
                            {
                                string cellDateString = cell.OldValue.GetDateString();
                                if (!string.IsNullOrWhiteSpace(cellDateString))
                                {
                                    cell.NewValue = GetNextMaxDateHeadCellValue(new List<ReplaceCell> { currentTableDateReplaceMatchItems.FirstOrDefault(), nextTableDateReplaceMatchItems.FirstOrDefault() }, cell.OldValue);
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                        else
                        {
                            //数据行 替换数据
                            var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                            if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue))
                            {
                                var mapDataRowIndex = nextTableDateRowCellList.FirstOrDefault(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue)?.StartRowIndex ?? -1;
                                if (mapDataRowIndex > 1)
                                {
                                    //从第二列开始替换数据
                                    var dataRowCellList = nextTableDateRowCellList.Where(w => w.StartRowIndex == mapDataRowIndex).ToList();
                                    foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                                    {
                                        var newCellValue = dataRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                        if (cell.OldValue != newCellValue)
                                        {
                                            cell.NewValue = newCellValue;
                                            cell.IsReplaceValue = true;
                                        }

                                    }
                                }

                            }
                        }

                    }

                }
            }
            #endregion

            #region 关键字
            if (currentTableKeywordReplaceMatchItems.Any() && nextTableKeywordReplaceMatchItems.Any())
            {
                var currentTableMatchItem = currentTableKeywordReplaceMatchItems.FirstOrDefault().ReplaceMatchItem;
                var nextTableMatchItem = currentTableKeywordReplaceMatchItems.FirstOrDefault().ReplaceMatchItem;
                var replaceItemList = WordTableConfigHelper.GetCellReplaceItemConfig();
                if (replaceItemList.Any(w => w.Key == currentTableMatchItem))
                {
                    //当前表格匹配项是key 从上往下替换
                    //当前表格匹配日期大于下一个表格匹配日期 从上往下替换
                    //清空当前表格 替换下一个表格
                    foreach (var row in table.Rows.Where(w => !w.IsHeadRow))
                    {
                        //数据行 清空数据
                        //判断当前表格数据行第一列内容是否存在于下一个表格数据行第一列单元格中 
                        var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                        if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue)
                            && nextTableDateRowCellList.Any(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue))
                        {
                            //清空从第二列开始数据
                            foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                            {
                                cell.NewValue = "";
                                cell.IsReplaceValue = true;
                            }
                        }
                    }
                    foreach (var row in nextWordTable.Rows.Where(w => !w.IsHeadRow))
                    {
                        //数据行 替换数据
                        var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                        if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue))
                        {
                            var mapDataRowIndex = tableDateRowCellList.FirstOrDefault(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue)?.StartRowIndex ?? -1;
                            if (mapDataRowIndex > 1)
                            {
                                //从第二列开始替换数据
                                var dataRowCellList = tableDateRowCellList.Where(w => w.StartRowIndex == mapDataRowIndex).ToList();
                                foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                                {
                                    cell.NewValue = dataRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                    cell.IsReplaceValue = true;
                                }
                            }

                        }
                    }
                }
                else
                {
                    //当前表格匹配项是value 从下往上替换
                    //当前表格匹配日期小于下一个表格匹配日期 从下往上替换
                    //清空下一个表格 替换当前表格
                    foreach (var row in nextWordTable.Rows.Where(w => !w.IsHeadRow))
                    {
                        //数据行 清空数据
                        var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                        if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue)
                            && tableDateRowCellList.Any(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue))
                        {
                            //清空从第二列开始数据
                            foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                            {
                                cell.NewValue = "";
                                cell.IsReplaceValue = true;
                            }
                        }
                    }
                    foreach (var row in table.Rows.Where(w => !w.IsHeadRow))
                    {
                        //数据行 替换数据
                        var rowFirstCell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == 1);
                        if (rowFirstCell != null && !string.IsNullOrWhiteSpace(rowFirstCell.OldValue))
                        {
                            var mapDataRowIndex = nextTableDateRowCellList.FirstOrDefault(w => w.StartColumnIndex == 1 && w.OldValue == rowFirstCell.OldValue)?.StartRowIndex ?? -1;
                            if (mapDataRowIndex > 1)
                            {
                                //从第二列开始替换数据
                                var dataRowCellList = nextTableDateRowCellList.Where(w => w.StartRowIndex == mapDataRowIndex).ToList();
                                foreach (var cell in row.RowCells.Where(w => w.StartColumnIndex > 1))
                                {
                                    cell.NewValue = dataRowCellList.FirstOrDefault(w => w.StartColumnIndex == cell.StartColumnIndex)?.OldValue;
                                    cell.IsReplaceValue = true;
                                }
                            }
                        }
                    }
                }
            }
            #endregion

        }

        /// <summary>
        /// 获取下一个日期
        /// </summary>
        /// <param name="replaceCells"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private static string GetNextMaxDateHeadCellValue(List<ReplaceCell> replaceCells, string cellValue)
        {
            var date = Convert.ToDateTime(cellValue.GetDateString());
            DateTime? nextMaxDate = null;
            if (replaceCells.Any(w => w.ReplaceMatchItemDate.Value.Month == 6))
            {
                //所有匹配项存在6月，代表当前表格是季度表
                //季度报
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

        /// <summary>
        /// 补充单元格
        /// </summary>
        /// <param name="rows">表格所有行</param>
        /// <param name="topFiveParagraphs">前五个段落</param>
        /// <returns></returns>
        private static void SupplementCell(WordTable table, List<WordParagraph> topFiveParagraphs)
        {
            //补全数据行单元格
            int maxDataRowCellCount = table.DataRows.Max(m => m.RowCells.Count);
            var lessThanMaxDataRowCellCountAllRowList = table.DataRows.Where(w => w.RowCells.Count < maxDataRowCellCount).ToList();
            if (lessThanMaxDataRowCellCountAllRowList.Any())
            {
                foreach (var row in lessThanMaxDataRowCellCountAllRowList)
                {
                    for (int cellIndex = 1; cellIndex <= maxDataRowCellCount; cellIndex++)
                    {
                        var cell = row.RowCells.FirstOrDefault(w => w.StartColumnIndex == cellIndex);
                        if (cell == null)
                        {
                            cell = new WordTableCell
                            {
                                StartRowIndex = row.RowNumber,
                                StartColumnIndex = cellIndex,
                                OldValue = "",
                            };
                            row.RowCells.Add(cell);
                        }
                    }
                    row.RowCells = row.RowCells.OrderBy(o => o.StartColumnIndex).ToList();
                }

            }

            //补全表头
            if (!table.HeadRows.Any())
            {
                //从下往上验证是否是表头段落
                topFiveParagraphs.Reverse();
                var headRowParagraphList = new List<WordParagraph>();

                foreach (var paragraph in topFiveParagraphs)
                {
                    Range paragraphRange = paragraph.Range;
                    var rangeText = paragraphRange.Text;

                    if (!rangeText.Contains("\t"))
                    {
                        //段落不包含\t 
                        break;
                    }
                    if (!string.IsNullOrWhiteSpace(rangeText.MatchWordTitle()) && rangeText.Count() < 50)
                    {
                        //段落是标题
                        break;
                    }
                    if (Regex.IsMatch(rangeText, @"[0-9]{3},"))
                    {
                        //表头不包含三位数，
                        break;
                    }
                    headRowParagraphList.Add(paragraph);
                }

                if (headRowParagraphList.Any())
                {
                    var paragraphSplitResultList = headRowParagraphList.OrderBy(o => o.ParagraphNumber).Select(s => s.Range.Text.Split('\t')).ToList();
                    if (paragraphSplitResultList.All(w => w.Count() == maxDataRowCellCount))
                    {
                        //分割后 所有表头行单元格数量都与最大数据行单元格数量一致 代表表头有效
                        var tableRowList = new List<WordTableRow>();
                        for (int rowIndex = 0; rowIndex < paragraphSplitResultList.Count; rowIndex++)
                        {
                            var currentRowSplitResult = paragraphSplitResultList[rowIndex];
                            var row = new WordTableRow()
                            {
                                RowNumber = rowIndex + 1,
                            };
                            for (int cellIndex = 0; cellIndex < currentRowSplitResult.Count(); cellIndex++)
                            {
                                row.RowCells.Add(new WordTableCell
                                {
                                    OldValue = currentRowSplitResult[cellIndex].RemoveSpaceAndEscapeCharacter(),
                                    StartRowIndex = rowIndex,
                                    StartColumnIndex = cellIndex,
                                    IsHeadColumn = true,
                                });
                            }
                            tableRowList.Add(row);
                        }

                        //重新加入数据行
                        foreach (var row in table.DataRows)
                        {
                            //重新计算数据行的行数和数据行单元格所在行数
                            row.RowNumber = tableRowList.Count + 1;
                            foreach (var cell in row.RowCells)
                            {
                                cell.StartRowIndex = row.RowNumber;
                            }
                            tableRowList.Add(row);
                        }

                        table.Rows = tableRowList;
                        table.TableContentStartParagraphNumber = headRowParagraphList.Min(w => w.ParagraphNumber);
                    }
                }
            }

            //lxz 2024-07-03 判断表格是否有人民币元行，和人民币行下一行是否空行；如果是则添加到表头
            SupplementRMBHeader(table);
        }

        /// <summary>
        /// 判断表格是否有人民币元行，和人民币行下一行是否空行；如果是则添加到表头
        /// </summary>
        /// <param name="table"></param>
        private static void SupplementRMBHeader(WordTable table)
        {
            //lxz判断表头是否有人民币
            var maxRowNumber = table.Rows.Max(x => x.RowNumber);
            var rowCount = table.Rows.Count;
            var headMaxNumber = 0;
            if (table.HeadRows.Any())
            {
                var lastrow = table.HeadRows.LastOrDefault();
                headMaxNumber = lastrow.RowNumber;
                //判断最后一行下一行，是否为空行；
                if (headMaxNumber + 1 <= maxRowNumber)
                {
                    var _row = table.Rows.Where(x => x.RowNumber == headMaxNumber + 1).FirstOrDefault();
                    if (_row != null && Regex.Replace(_row.RowContent, @"\s", "") == "")
                    {
                        _row.RowCells.ForEach(c => { c.IsHeadColumn = true; });
                        lastrow = _row;
                        headMaxNumber = headMaxNumber + 1;
                    }
                }
            }
            var row = table.Rows.Where(r => r.RowCells.Any(x => x.StartColumnIndex != 1 && !Regex.IsMatch(x.OldValue, @"(\d{1,3},\d+)+|\d{1,3}") && (x.OldValue.Trim().Equals("人民币") || x.OldValue.Trim().Equals("折合人民币元") || x.OldValue.Trim().Equals("人民币元") || x.OldValue.Trim().Equals("%") || x.OldValue.Trim().Equals("美元")))).FirstOrDefault();
            if (row != null)
            {
                if (headMaxNumber < row.RowNumber)
                {
                    headMaxNumber = row.RowNumber;
                }
                var nextRow = table.Rows.Where(x => x.RowNumber == headMaxNumber + 1).FirstOrDefault();
                if (nextRow != null && Regex.Replace(nextRow.RowContent, @"\s", "") == "")
                {
                    headMaxNumber = nextRow.RowNumber;
                }
            }
            if (headMaxNumber > 0)
            {
                foreach (var item in table.Rows)
                {
                    if (item.RowNumber <= headMaxNumber)
                    {
                        item.RowCells.ForEach(c => { c.IsHeadColumn = true; });
                        //table.HeadRows.Add(item);
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// 匹配制表位表格单元格Range
        /// </summary>
        /// <param name="table"></param>
        private static void MatchTabStopTableCellRange(WordTable table)
        {
            foreach (WordTableRow row in table.Rows)
            {
                //根据未使用段落计算出来的表格不需要计算行Range
                if (table.TableSourceType != TableSourceTypeEnum.TabStopCompute)
                {
                    #region 使用编辑距离判断相似都最高的段落

                    //lxz 2024-07-11 使用编辑距离判断相似都最高的段落
                    //Range rowRange = table.ContentParagraphs.Where(w => w.Text.Contains(row.RowContent.RemoveWordTitle())).FirstOrDefault()?.Range;
                    Range rowRange = null;
                    //lxz 2024-07-11 使用编辑距离来判断文本相似度，取相似度最高的段落 Levenshtein_Distance
                    var tarray = table.ContentParagraphs.Where(w => w.Text.Contains(row.RowContent.RemoveWordTitle())).ToArray();
                    if (tarray.Any())
                    {
                        if (tarray.Length == 1)
                        {
                            rowRange = tarray.First().Range;
                        }
                        else
                        {
                            List<(WordParagraph paragraph, double val)> tList = new List<(WordParagraph paragraph, double val)>();

                            foreach (var item in tarray)
                            {
                                var val = StringHelper.Levenshtein_Distance(row.RowContent, item.Text);
                                tList.Add((item, val));
                            }
                            rowRange = tList.OrderByDescending(x => x.val).First().paragraph.Range;
                        }
                    }

                    if (rowRange == null)
                    {
                        string errorMsg = $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})第{row.RowNumber}行({row.RowContent})未能匹配到Word段落!";
                        table.ErrorMsgs.Add(errorMsg);
                        table.OperationType = OperationTypeEnum.ConsoleError;
                        errorMsg.Console(ConsoleColor.Red);
                        break;
                    }
                    row.IsMatchRowRange = true;
                    row.Range = rowRange;
                    #endregion

                }
                string rowRangeText = row.Range.Text;
                int lastColumnIndex = row.RowCells.LastOrDefault().StartColumnIndex;
                int nextCellStartIndex = 0;
                foreach (var cell in row.RowCells)
                {
                    string cellIndexValue = string.Empty;
                    if (cell.StartColumnIndex != lastColumnIndex)
                    {
                        cellIndexValue = cell.OldValue + "\t";
                    }
                    else
                    {
                        cellIndexValue = cell.OldValue + "\r";
                    }

                    Range cellRange = row.Range.Duplicate;
                    int? cellStartIndex = null;
                    if (string.IsNullOrWhiteSpace(cell.OldValue))
                    {
                        if (cell.StartColumnIndex != lastColumnIndex)
                        {
                            //空单元格用\t定位Range
                            cellStartIndex = rowRangeText.IndexOf("\t", nextCellStartIndex);
                        }
                        else
                        {
                            //最后一个空单元格用\r定位Range
                            cellStartIndex = rowRangeText.IndexOf("\r", nextCellStartIndex);
                        }
                    }
                    else
                    {
                        cellStartIndex = rowRangeText.IndexOf(cell.OldValue, nextCellStartIndex);
                    }

                    int cellEndIndex = cellStartIndex.Value + cellIndexValue.Length;
                    int moveNumber = rowRangeText.Length - cellEndIndex;
                    cellRange.MoveStart(WdUnits.wdCharacter, cellStartIndex);
                    cellRange.MoveEnd(WdUnits.wdCharacter, -moveNumber);
                    cell.Range = cellRange;
                    nextCellStartIndex = cellEndIndex;
                    string moveCellValue = cell.Range.Text;

                }
            }
        }



        /// <summary>
        /// 生成制表位表格单元格新值
        /// </summary>
        /// <param name="tables"></param>
        private static void BuildTabStopTableCellNewValue(List<WordTable> tables, Action<object, Events.NodeNotifyEventArgs> errorMsg_Event = null)
        {
            var replaceItemList = WordTableConfigHelper.GetCellReplaceItemConfig();
            int lastTableIndex = tables.IndexOf(tables.LastOrDefault());
            for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
            {
                string errorMsg = string.Empty;
                var table = tables[tableIndex];
                if (table.IsTabStopTable)
                {
                    try
                    {
                        if (table.OperationType == OperationTypeEnum.ConsoleError || table.OperationType == OperationTypeEnum.ChangeColor)
                        {
                            continue;
                        }
                        if (!table.IsMatchWordParagraph || table.Rows.Any(w => !w.IsMatchRowRange))
                        {
                            errorMsg = $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})未能匹配到Word段落范围";
                            table.OperationType = OperationTypeEnum.ChangeColor;
                            table.ErrorMsgs.Add(errorMsg);
                            errorMsg.Console(ConsoleColor.Red);
                            continue;
                        }
                        if (!table.HeadRows.Any())
                        {
                            errorMsg = $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})未能识别到表头";
                            table.OperationType = OperationTypeEnum.ChangeColor;
                            table.ErrorMsgs.Add(errorMsg);
                            errorMsg.Console(ConsoleColor.Red);
                            continue;
                        }

                        #region 同表左右替换 判断当前表格所有表头是否包含两个及以上不同日期或者包含任意一组关键字

                        var horizontalHeadRowCellList = GetHorizontalMergeTableHeadRow(table.HeadRows);

                        var horizontalDateReplaceMatchItemList = horizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                        && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date).ToList();
                        var horizontalDateReplaceMatchItemGroupCount = horizontalDateReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                        var horizontalKeywordReplaceMatchItemList = horizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                        && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Keyword).ToList();
                        var filterHorizontalKeywordReplaceMatchItemList = new List<ReplaceCell>();
                        horizontalKeywordReplaceMatchItemList.ForEach(matchItem =>
                        {
                            var matchItemKeyvaluePair = replaceItemList.FirstOrDefault(w => w.Key == matchItem.ReplaceMatchItem || w.Value == matchItem.ReplaceMatchItem);
                            bool isIncludeKeyvaluePair = new string[] { matchItemKeyvaluePair.Key, matchItemKeyvaluePair.Value }.All(w => horizontalKeywordReplaceMatchItemList.Select(s => s.ReplaceMatchItem).Contains(w));
                            if (isIncludeKeyvaluePair && matchItem.ReplaceMatchItemType != ReplaceMatchItemTypeEnum.Disturb)
                            {
                                filterHorizontalKeywordReplaceMatchItemList.Add(matchItem);
                            }
                        });
                        horizontalKeywordReplaceMatchItemList = filterHorizontalKeywordReplaceMatchItemList;

                        var horizontalKeywordReplaceMatchItemGroupCount = horizontalKeywordReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                        if (horizontalDateReplaceMatchItemGroupCount >= 2 ||
                           horizontalKeywordReplaceMatchItemGroupCount >= 2)
                        {
                            //执行同表跨列替换逻辑
                            SameTableCrossColumnReplace(table, horizontalDateReplaceMatchItemList, horizontalKeywordReplaceMatchItemList);
                            table.OperationType = OperationTypeEnum.ReplaceText;

                            //lxz 2024-07-22 添加逻辑
                            ReSetTableOperationType(table, horizontalHeadRowCellList, null);
                            continue;
                        }
                        #endregion

                        #region 同表上下替换 判断当前表格第一列是否包含两个及以上不同日期或者包含任意一组关键字 且上下表
                        var verticalHeadRowCellList = GetVerticalTableHeadRow(table.Rows);
                        var isTheSameHeadRow = IsTheSameHeadRow(table, tables);
                        if (!isTheSameHeadRow)
                        {
                            var verticalDateReplaceMatchItemList = verticalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                            && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date).ToList();
                            var verticalDateReplaceMatchItemGroupCount = verticalDateReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                            var verticalKeywordReplaceMatchItemList = verticalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                            && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Keyword).ToList();
                            var filterVerticalKeywordReplaceMatchItemList = new List<ReplaceCell>();
                            verticalKeywordReplaceMatchItemList.ForEach(matchItem =>
                            {
                                var matchItemKeyvaluePair = replaceItemList.FirstOrDefault(w => w.Key == matchItem.ReplaceMatchItem || w.Value == matchItem.ReplaceMatchItem);
                                bool isIncludeKeyvaluePair = new string[] { matchItemKeyvaluePair.Key, matchItemKeyvaluePair.Value }.All(w => verticalKeywordReplaceMatchItemList.Select(s => s.ReplaceMatchItem).Contains(w));
                                if (isIncludeKeyvaluePair && matchItem.ReplaceMatchItemType != ReplaceMatchItemTypeEnum.Disturb)
                                {
                                    filterVerticalKeywordReplaceMatchItemList.Add(matchItem);
                                }
                            });
                            verticalKeywordReplaceMatchItemList = filterVerticalKeywordReplaceMatchItemList;
                            var verticalKeywordReplaceMatchItemGroupCount = verticalKeywordReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                            if (verticalDateReplaceMatchItemGroupCount >= 2 ||
                               verticalKeywordReplaceMatchItemGroupCount >= 2)
                            {
                                //执行同表跨行替换逻辑
                                SameTableCrossRowReplace(table, verticalDateReplaceMatchItemList, verticalKeywordReplaceMatchItemList);
                                table.OperationType = OperationTypeEnum.ReplaceText;

                                //lxz 2024-07-22 添加逻辑
                                ReSetTableOperationType(table, horizontalHeadRowCellList, verticalHeadRowCellList);
                                continue;
                            }
                        }

                        #endregion

                        #region 跨表上下替换 只有当前表非最后一个表格且与下一个表格表头除了匹配项完全一致
                        if (tableIndex < lastTableIndex)
                        {
                            var nextTable = tables[tableIndex + 1];

                            //判断下一个表是否符合替换条件 
                            if (nextTable.IsMatchWordParagraph && nextTable.Rows.All(w => w.IsMatchRowRange))
                            {
                                //判断当前表格与下一个表格是否表头除开日期部分是否完全一致 上下两个表格均有一个匹配项
                                //且当前表格第一列所有行中内容是否存在任意一项在下一个表格第一列所有行中存在
                                var nextTableHorizontalHeadRowCellList = nextTable != null ? GetHorizontalMergeTableHeadRow(nextTable.HeadRows) : new List<ReplaceCell>();

                                var nextTableHorizontalDateReplaceMatchItemList = nextTableHorizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                                 && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Date).ToList();
                                var nextTableHorizontalDateReplaceMatchItemGroupCount = nextTableHorizontalDateReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                                var nextTableHorizontalKeywordReplaceMatchItemList = nextTableHorizontalHeadRowCellList.Where(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)
                                  && w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Keyword).ToList();
                                var nextTableHorizontalKeywordReplaceMatchItemGroupCount = nextTableHorizontalKeywordReplaceMatchItemList.GroupBy(g => g.ReplaceMatchItem).Count();

                                string tableHeadRowContent = string.Join("", horizontalHeadRowCellList.Select(s => s.CellValue)).ReplaceDate();
                                string nextTableHeadRowContent = string.Join("", nextTableHorizontalHeadRowCellList.Select(s => s.CellValue)).ReplaceDate();

                                string currentTableFirstReplaceMatchItem = horizontalKeywordReplaceMatchItemList.FirstOrDefault()?.ReplaceMatchItem;
                                string nextTableFirstReplaceMatchItem = nextTableHorizontalKeywordReplaceMatchItemList.FirstOrDefault()?.ReplaceMatchItem;
                                bool isKeyValuePair = !string.IsNullOrWhiteSpace(currentTableFirstReplaceMatchItem) && !string.IsNullOrWhiteSpace(nextTableFirstReplaceMatchItem)
                                    && replaceItemList.Count(w => (w.Key == currentTableFirstReplaceMatchItem && w.Value == nextTableFirstReplaceMatchItem) ||
                                    (w.Key == nextTableFirstReplaceMatchItem && w.Value == currentTableFirstReplaceMatchItem)) == 1;

                                if (tableHeadRowContent == nextTableHeadRowContent //上下两个表除开日期部分表头一致
                                    && ((horizontalDateReplaceMatchItemGroupCount == 1 && nextTableHorizontalDateReplaceMatchItemGroupCount == 1)//上下两个表都有一个日期匹配项
                                    || (horizontalKeywordReplaceMatchItemGroupCount == 1 && nextTableHorizontalKeywordReplaceMatchItemGroupCount == 1 && isKeyValuePair))) //上下两个表都有一个关键字匹配项，且是一堆键值对
                                {
                                    //执行跨表替换逻辑
                                    CrossTableReplace(table, horizontalDateReplaceMatchItemList, horizontalKeywordReplaceMatchItemList,
                                        nextTable, nextTableHorizontalDateReplaceMatchItemList, nextTableHorizontalKeywordReplaceMatchItemList);
                                    table.OperationType = OperationTypeEnum.ReplaceText;
                                    nextTable.OperationType = OperationTypeEnum.ReplaceText;
                                    //循环跳过下一个表
                                    tableIndex++;

                                }
                            }
                        }

                        #endregion

                        //lxz 2024-07-01 添加逻辑
                        ReSetTableOperationType(table, horizontalHeadRowCellList, verticalHeadRowCellList);
                    }
                    catch (Exception ex)
                    {
                        errorMsg = $"第{table.PageNumber}页第{table.TableNumber}个表格({table.FirstRowContent})生成单元格新值失败，{ex.Message}";
                        table.OperationType = OperationTypeEnum.ChangeColor;
                        table.ErrorMsgs.Add(errorMsg);
                        errorMsg.Console(ConsoleColor.Red);
                        ex.StackTrace.Console(ConsoleColor.Red);
                        if (errorMsg_Event != null)
                        {
                            errorMsg_Event(null, new Events.NodeNotifyEventArgs
                            {
                                Type = 2,
                                Message = $"{errorMsg}"
                            });
                        }
                    }
                }

            }
        }

        /// <summary>
        /// 重新判断设置表格是否替换表头颜色
        /// </summary>
        /// <param name="table"></param>
        /// <param name="horizontalHeadRowCellList"></param>
        /// <param name="verticalHeadRowCellList"></param>
        private static void ReSetTableOperationType(WordTable table, List<ReplaceCell> horizontalHeadRowCellList, List<ReplaceCell> verticalHeadRowCellList)
        {
            //lxz 2024-07-01 添加逻辑
            //表格表头包含年份，却没有执行上面的替换逻辑，则表头应该替换颜色
            if (table.OperationType == OperationTypeEnum.NotOperation && (
              horizontalHeadRowCellList.Any(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem))
              || (verticalHeadRowCellList != null && verticalHeadRowCellList.Any(w => !string.IsNullOrWhiteSpace(w.ReplaceMatchItem)))
              || horizontalHeadRowCellList.Any(w => w.ReplaceMatchItemType == ReplaceMatchItemTypeEnum.Disturb)
              ))
            {
                table.OperationType = OperationTypeEnum.ChangeColor;
            }
            else if (table.OperationType == OperationTypeEnum.ReplaceText && !table.Rows.Where(x => x.RowCells.Any(c => c.IsReplaceValue)).Any())
            {
                table.OperationType = OperationTypeEnum.ChangeColor;
            }
        }
        #endregion

        #region Common

        /// <summary>
        /// 判断当前表格与前一个和下一个是否相同表头行
        /// </summary>
        /// <param name="currentTable"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        private static bool IsTheSameHeadRow(WordTable currentTable, List<WordTable> tables)
        {
            var currentTableHeadRowContent = string.Join("", GetHorizontalMergeTableHeadRow(currentTable.HeadRows).Select(s => s.CellValue)).ReplaceDate();
            var prevTableNumber = currentTable.TableNumber - 1;
            var prevTable = tables.FirstOrDefault(w => w.TableNumber == prevTableNumber);
            var prevTableHeadRowContent = string.Empty;
            if (prevTable != null)
            {
                prevTableHeadRowContent = string.Join("", GetHorizontalMergeTableHeadRow(prevTable.HeadRows).Select(s => s.CellValue)).ReplaceDate();
            }
            var nextTableNumber = currentTable.TableNumber + 1;
            var nextTable = tables.FirstOrDefault(w => w.TableNumber == nextTableNumber);
            string nextTableHeadRowContent = string.Empty;
            if (nextTable != null)
            {
                nextTableHeadRowContent = string.Join("", GetHorizontalMergeTableHeadRow(nextTable.HeadRows).Select(s => s.CellValue)).ReplaceDate();
            }
            bool isTheSameHeadRow = currentTableHeadRowContent == prevTableHeadRowContent || currentTableHeadRowContent == nextTableHeadRowContent;
            return isTheSameHeadRow;
        }

        /// <summary>
        /// 合并表格
        /// </summary>
        /// <param name="tables"></param>
        private static void MergeTables(List<WordTable> tables, Action<object, Events.NodeNotifyEventArgs> errorMsg_Event)
        {
            /*
             * 合并表格，第一个表格有表头，后续连续表格无表头且列数相同，则认为是同一个表格需要合并为一个表格；
             * 1.判断列数相同
             * 2.判断子表无表头
             * 3.存储连续无表头子表
             * 4.合并子表
             * 5.删除合并后的子表，并重新设置表格编号
             */
            WordTable currnetTable = null;
            try
            {
                var tableCount = tables.Count;
                int index = 0;
                //需要删除的表格下标
                var removeList = new List<int>();

                //根据下标循环表格
                while (index < tableCount)
                {
                    var table = tables[index];
                    currnetTable = table;
                    var nextIndex = index + 1;
                    //获取当前表格的最大列表数
                    var tableCol = table.Rows.Max(x => x.RowCells.Count);

                    //默认数据开始列从第一列开始，下标从0开始
                    var startNumCol = 1;
                    //获取数据开始列
                    for (int _colIndex = 1; _colIndex < tableCol; _colIndex++)
                    {
                        var isEmpty = true;
                        var isBreak = false;
                        for (int _rowIndex = 0; _rowIndex < table.Rows.Count; _rowIndex++)
                        {
                            var _row = table.Rows[_rowIndex];
                            if (_row.RowCells.Count <= _colIndex)
                            {
                                continue;
                            }
                            var _cell = _row.RowCells[_colIndex];
                            if (!string.IsNullOrWhiteSpace(_cell.OldValue) && isEmpty)
                            {
                                isEmpty = false;
                            }
                            if (Regex.IsMatch(_cell.OldValue, @"\d{4}|[\u4e00-\u9fa5]"))
                            {
                                continue;
                            }
                            if (Regex.IsMatch(_cell.OldValue, @"\d+"))
                            {
                                startNumCol = _colIndex;
                                isBreak = true;
                                break;
                            }
                        }
                        if (isEmpty || isBreak)
                        {
                            startNumCol = _colIndex;
                            break;
                        }

                    }

                    //当前表格子级表格
                    var currentTableList = new List<WordTable>();
                    while (nextIndex < tableCount)
                    {
                        var nextTable = tables[nextIndex];
                        var nextTableCol = nextTable.Rows.Max(x => x.RowCells.Count);

                        if (nextTableCol == tableCol)
                        {
                            var isEmpty = true;
                            if (nextTable.HeadRows.Any())
                            {
                                for (int i = 0; i < nextTable.HeadRows.Count; i++)
                                {
                                    var nextTableHeadRow = nextTable.HeadRows[i];
                                    for (int c = startNumCol; c < nextTableHeadRow.RowCells.Count; c++)
                                    {
                                        var val = nextTableHeadRow.RowCells[c]?.OldValue?.Trim();
                                        if (!string.IsNullOrWhiteSpace(val) && Regex.IsMatch(val, @"\d{4}|[\u4e00-\u9fa5]|[a-zA-Z]"))
                                        {
                                            isEmpty = false;
                                            break;
                                        }
                                    }
                                    if (!isEmpty)
                                    {
                                        break;
                                    }
                                }
                            }
                            if (!isEmpty)
                            {
                                break;
                            }
                            currentTableList.Add(nextTable);
                            removeList.Add(nextIndex);
                            index = nextIndex;
                            nextIndex = index + 1;
                        }
                        else
                        {
                            break;
                        }
                    }
                    //合并表格
                    if (currentTableList.Any())
                    {
                        var tempRowNumber = table.Rows.Max(x => x.RowNumber);
                        foreach (var subTable in currentTableList)
                        {
                            List<WordTableRow> tRows = new List<WordTableRow>();
                            for (int idx = 0; idx < subTable.Rows.Count; idx++)
                            {
                                tempRowNumber++;
                                var subRow = subTable.Rows[idx];
                                WordTableRow trow = new WordTableRow()
                                {
                                    IsMatchRowRange = subRow.IsMatchRowRange,
                                    Range = subRow.Range,
                                    RowNumber = tempRowNumber,
                                };

                                for (int c = 0; c < subRow.RowCells.Count; c++)
                                {
                                    var subCell = subRow.RowCells[c];
                                    WordTableCell tcell = new WordTableCell
                                    {
                                        ActualStartColumnIndex = subCell.ActualStartColumnIndex,
                                        CellChars = subCell.CellChars,
                                        ChilderCells = subCell.ChilderCells,
                                        ColSpan = subCell.ColSpan,
                                        IsHeadColumn = false,
                                        IsReplaceValue = subCell.IsReplaceValue,
                                        Length = subCell.Length,
                                        MinX = subCell.MinX,
                                        MinY = subCell.MinY,
                                        NewValue = subCell.NewValue,
                                        Offset = subCell.Offset,
                                        OldValue = subCell.OldValue,
                                        PageNumber = subCell.PageNumber,
                                        Range = subCell.Range,
                                        RowSpan = subCell.RowSpan,
                                        VMergeVal = subCell.VMergeVal,
                                        YPositiondifference = subCell.YPositiondifference,
                                        StartColumnIndex = subCell.StartColumnIndex,
                                        StartRowIndex = trow.RowNumber,
                                    };
                                    trow.RowCells.Add(tcell);
                                }
                                tRows.Add(trow);
                            }
                            table.Rows.AddRange(tRows);
                        }
                    }
                    index++;
                }
                if (removeList.Any())
                {
                    removeList.Reverse();
                    foreach (var item in removeList)
                    {
                        tables.RemoveAt(item);
                    }
                    var t_idx = 1;
                    foreach (var item in tables)
                    {
                        item.TableNumber = t_idx;

                        t_idx++;
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "";
                if (currnetTable != null)
                {
                    //currnetTable.PageNumber;
                    msg = $",开始行内容:{currnetTable.FirstRowContent};结束行内容：{currnetTable.LastRowContent}";
                }
                //throw new Exception($"合并表格失败{msg}！ex:{ex.Message}", ex);
                if (errorMsg_Event != null)
                {
                    errorMsg_Event(null, new Events.NodeNotifyEventArgs
                    {
                        Type = 2,
                        Message = $"合并表格失败{msg}！ex:{ex.Message}"
                    });
                }
            }
        }

        /// <summary>
        /// 拆分表格
        /// </summary>
        /// <param name="tables"></param>
        private static List<WordTable> SplitTables(List<WordTable> tables, Action<object, Events.NodeNotifyEventArgs> errorMsg_Event = null)
        {
            try
            {
                var needAddTableList = new List<WordTable>();
                var needDeleteTableList = new List<WordTable>();
                foreach (var table in tables)
                {
                    var splitResultList = FindTables(table);
                    if (splitResultList.Count > 1)
                    {
                        needAddTableList.AddRange(splitResultList);
                        needDeleteTableList.Add(table);
                    }
                }
                if (needAddTableList.Any() && needDeleteTableList.Any())
                {
                    foreach (var deleteTable in needDeleteTableList)
                    {
                        tables.Remove(deleteTable);
                    }
                    tables.AddRange(needAddTableList);
                    tables = tables.OrderBy(o => o.PageNumber).ToList();
                    tables.ForEach(table =>
                    {
                        table.TableNumber = tables.IndexOf(table) + 1;
                    });
                }
            }
            catch (Exception ex)
            {
                if (errorMsg_Event != null)
                {
                    errorMsg_Event(null, new Events.NodeNotifyEventArgs
                    {
                        Type = 2,
                        Message = $"制表位表格拆分错误，ex:{ex.Message}"
                    });
                }

            }
            return tables;
        }

        #endregion

        #region 识别段落中的制表位表格

        /// <summary>
        /// 根据连续段落获取制表位表格
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <returns></returns>
        private static List<WordTable> FindTables(List<WordParagraph> paragraphs)
        {
            var tableList = new List<WordTable>();

            var tableFirst = false;
            var tableRowStart = -1;
            List<int> c_rows = new List<int>();
            List<int> s_rows = new List<int>();
            List<int> e_rows = new List<int>();

            // find currency row
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var currentParagraph = paragraphs[i];
                var text = currentParagraph.OldText;
                var row = text.Split('\t');
                if (row.Length > 1 && isCurrencyRow(row) && row.LastOrDefault().Length < 20)
                {
                    c_rows.Add(i);
                    if (tableFirst == false && i < tableRowStart)
                    {
                        tableFirst = true;
                    }
                }
            }

            // 没有找到的话说明没有人民币行，直接根据tab去找
            if (c_rows.Count == 0)
            {
                // 如果某一行为空，且上下两行都有数据，并且tab数量相差不超过1，这行当做货币行
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var t_para = paragraphs[i];

                    if (i > 0 && i < paragraphs.Count - 1 && isBlank(t_para))
                    {

                        var prePara = paragraphs[i - 1];
                        var nextPara = paragraphs[i + 1];

                        int preRowTabCount = getTabStopsCount(prePara, true);//WordUtil.getTabStops(paras.get(i - 1), true).size();
                        int nextRowTabCount = getTabStopsCount(nextPara, true);//WordUtil.getTabStops(paras.get(i + 1), true).size();

                        //lxz 2021-11-判断表格列数应该大于1，一个表格至少两列，
                        if (Math.Abs(preRowTabCount - nextRowTabCount) <= 1 && nextRowTabCount > 1 && preRowTabCount > 1)
                        {
                            //如果当前tabCount=1时可能为段落,增加长度小于等于50判断				zqb		2021-10-08 22:38
                            //if (preRowTabCount != 1 && nextRowTabCount != 1 && (paras.get(i - 1).getText().length() <= 50 && paras.get(i + 1).getText().length() <= 50))
                            if (preRowTabCount != 1 && nextRowTabCount != 1 && (prePara.Range.Text.Length <= 50 && nextPara.Range.Text.Length <= 50))
                            {
                                c_rows.Add(i - 1);
                                if (tableFirst == false && i < tableRowStart)
                                {
                                    tableFirst = true;
                                }
                            }
                        }
                    }
                }
            }

            // find table start
            for (int i = 0; i < c_rows.Count; i++)
            {
                int s_row = c_rows[i];
                while (s_row > 0)
                {
                    s_row--;

                    /**
                     * 人民币向上找，第一个空行就是start ，特殊情况： 1.表格最上面不是空行，可能是一句描述，此时判断有没有tab，没有的话也算结束
                     */
                    var s_para = paragraphs[s_row];
                    if (isBlank(s_para) || getTabStopsCount(s_para) == 0)
                    {
                        s_row++; // 找到的那一行的下一行是开始行
                        break;
                    }
                }

                if (s_row == -1)
                {
                    s_row = 0;
                }
                s_rows.Add(s_row);
            }

            // find table end
            for (int i = 0; i < c_rows.Count; i++)
            {
                int e_row = 0;
                if (i == c_rows.Count - 1)
                {
                    e_row = paragraphs.Count - 1;
                }
                else
                {
                    e_row = s_rows[i + 1] - 1;
                }
                int c_row = c_rows[i];

                /**
                 * 从人民币行下的空行向下找，一般表格的格式是 head+currecyRow+emptyRow+body,特殊场景：
                 * 1.currecny下面有时候还会有非空行，比如% 2.currency下面没有空行...
                 */
                int bodyStart = c_row;
                var bodyStartPara = paragraphs[bodyStart];
                while (bodyStart < e_row && !isBlank(bodyStartPara))
                {
                    // 目前还没有遇到过人民币下面超过2行还没开始表格内容的
                    if (bodyStart - c_row >= 2)
                    {
                        bodyStart = c_row;
                        bodyStartPara = paragraphs[bodyStart];
                        break;
                    }
                    bodyStart++;
                    bodyStartPara = paragraphs[bodyStart];
                }

                for (int j = bodyStart + 1; j <= e_row; j++)
                {
                    var t_para = paragraphs[j];
                    string matchWordTitle = t_para.OldText.MatchWordTitle();
                    if (t_para.OldText.Contains("\f"))
                    {
                        //遇到分页符当前表结束
                        e_row = j - 1;
                        break;
                    }
                    else if (!string.IsNullOrWhiteSpace(matchWordTitle) && t_para.OldText.Contains("\t"))
                    {
                        //遇到标题且带\t
                        e_row = j - 1;
                        break;
                    }
                    else if (isBlank(t_para))
                    {
                        /**
                         * 空行但是是表格内容的场景：
                         * 
                         * Scene 1.空行+表格内标题+内容
                         * --EmptyRow--
                         * and after crediting in other income:
                         * Interest income	8,289	9,297
                         * 
                         * Scene 2.空行+表格内标题+空行+内容
                         * --EmptyRow--
                         * and after crediting in other income:
                         * --EmptyRow--
                         * Interest income	8,289	9,297
                         */

                        if (j + 2 <= e_row)
                        {
                            var next = paragraphs[j + 1];
                            var nnext = paragraphs[j + 2];
                            if (!isBlank(next))
                            {
                                //Scene 1
                                if (!isBlank(nnext) && getTabStopsCount(nnext, true) > 0)
                                {
                                    j = j + 2;
                                    continue;
                                }
                                //Scene 2
                                if (j + 3 <= e_row)
                                {
                                    var nnnext = paragraphs[j + 3];
                                    if (isBlank(nnext) && getTabStopsCount(nnnext, true) > 0)
                                    {
                                        j = j + 3;
                                        continue;
                                    }
                                }
                            }
                        }

                        e_row = j - 1;
                        break;
                    }
                    else if (getTabStopsCount(t_para, true) == 0)
                    {
                        // 存在人民币行下面试空行，空行下只有一个汇总header的情况
                        if (j + 1 <= e_row)
                        {
                            var next = paragraphs[j + 1];
                            if (getTabStopsCount(next, true) > 0 && !isBlank(next))
                            {
                                j = j + 1;
                                continue;
                            }
                        }
                        e_row = j - 1;
                        break;
                    }
                }

                e_rows.Add(e_row);
            }

            var maxCol = 0;
            for (int i = 0; i < c_rows.Count; i++)
            {
                int start = s_rows[i];
                int currency = c_rows[i];
                int end = e_rows[i];
                var tableHeadRowParagraphList = new List<WordParagraph>();
                for (int k = start; k <= currency; k++)
                {
                    var para = paragraphs[k];
                    if (IsParaUnderline(para))
                    {
                        continue;
                    }
                    tableHeadRowParagraphList.Add(para);
                    var templegth = para.Range.Text.Split('\t').Length;
                    if (maxCol < templegth)
                    {
                        maxCol = templegth;
                    }
                }

                // -----------get table profile data
                var tableDataRowParagraphList = new List<WordParagraph>();
                for (int k = currency + 1; k <= end; k++)
                {
                    var para = paragraphs[k];
                    if (IsParaUnderline(para))
                    {
                        continue;
                    }
                    tableDataRowParagraphList.Add(para);
                    var templegth = para.OldText.Split('\t').Length;
                    if (maxCol < templegth)
                    {
                        maxCol = templegth;
                    }
                }

                var table = new WordTable()
                {
                    TableSourceType = TableSourceTypeEnum.TabStopCompute,
                    PageNumber = paragraphs[start].PageNumber,
                    TableNumber = tableList.Count + 1,
                    IsMatchWordParagraph = true,
                    TableContentStartParagraphNumber = tableHeadRowParagraphList.Where(w => !w.IsEmptyParagraph).Min(w => w.ParagraphNumber),
                    TableContentEndParagraphNumber = tableDataRowParagraphList.Where(w => !w.IsEmptyParagraph).Max(w => w.ParagraphNumber)
                };
                foreach (WordParagraph para in tableHeadRowParagraphList)
                {
                    if (para.IsEmptyParagraph)
                    {
                        continue;
                    }
                    //var paraSplitResults=para.OldText.StartsWith("\t")? para.OldText.Remove(0,1).TrimEnd('\r').Split('\t')
                    //    :para.OldText.TrimEnd('\r').Split('\t');
                    var paraSplitResults = para.OldText.TrimEnd('\r').Split('\t');
                    var tableRow = new WordTableRow
                    {
                        RowNumber = table.Rows.Count + 1,
                        Range = para.Range,
                        IsMatchRowRange = true
                    };
                    foreach (var paraSplitResult in paraSplitResults)
                    {
                        tableRow.RowCells.Add(new WordTableCell
                        {
                            StartRowIndex = tableRow.RowNumber,
                            StartColumnIndex = tableRow.RowCells.Count + 1,
                            OldValue = paraSplitResult.RemoveSpaceAndEscapeCharacter().ConvertCharToHalfWidth(),
                            IsHeadColumn = true
                        });
                    }
                    table.Rows.Add(tableRow);
                    table.ContentParagraphs.Add(para);
                }
                foreach (WordParagraph para in tableDataRowParagraphList)
                {
                    if (para.IsEmptyParagraph)
                    {
                        continue;
                    }
                    //var paraSplitResults = para.OldText.StartsWith("\t") ? para.OldText.Remove(0, 1).TrimEnd('\r').Split('\t')
                    //     : para.OldText.TrimEnd('\r').Split('\t');
                    var paraSplitResults = para.OldText.TrimEnd('\r').Split('\t');
                    var tableRow = new WordTableRow
                    {
                        RowNumber = table.Rows.Count + 1,
                        Range = para.Range,
                        IsMatchRowRange = true
                    };
                    foreach (var paraSplitResult in paraSplitResults)
                    {
                        tableRow.RowCells.Add(new WordTableCell
                        {
                            StartRowIndex = tableRow.RowNumber,
                            StartColumnIndex = tableRow.RowCells.Count + 1,
                            OldValue = paraSplitResult.RemoveSpaceAndEscapeCharacter().ConvertCharToHalfWidth()
                        });
                    }
                    table.Rows.Add(tableRow);
                    table.ContentParagraphs.Add(para);
                }
                tableList.Add(table);
            }

            return tableList;

        }

        /// <summary>
        /// 根据表格行获取拆分表格
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private static List<WordTable> FindTables(WordTable table)
        {
            var tableList = new List<WordTable>();

            var tableStartRowIndexList = new List<int>();
            var tableHeadRowEndRowIndexList = new List<int>();

            //查找人民币元当表头结束行
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                if (new string[] { "人民币", "人民币元" }.Any(w => row.RowContent.Contains(w))
                    && row.RowCells.Where(x => x.StartColumnIndex != 1 && x.OldValue != null)
                    .Any(x => Regex.IsMatch(x.OldValue, "人民币|人民币元") && !Regex.IsMatch(x.OldValue, @"(\d{1,3},\d+)+|\d{1,3}")))
                {
                    tableHeadRowEndRowIndexList.Add(i);
                }
            }

            //没有人民币元 查找上下两行都不为空的空行的上一行当表头结束行
            if (tableHeadRowEndRowIndexList.Count == 0)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (i > 0 && i < table.Rows.Count - 1 && string.IsNullOrWhiteSpace(table.Rows[i].RowContent.Trim()))
                    {
                        var prevRow = table.Rows[i - 1];
                        var nextRow = table.Rows[i + 1];
                        if (!string.IsNullOrWhiteSpace(prevRow.RowContent.Trim()) && !string.IsNullOrWhiteSpace(nextRow.RowContent.Trim()))
                        {
                            tableHeadRowEndRowIndexList.Add(i - 1);
                        }
                    }
                }
            }

            //如果当前表只找到一个表头 当前表不拆分
            if (tableHeadRowEndRowIndexList.Count <= 1)
            {
                return tableList;
            }

            //从表头结束行往上找 遇到空行和数据行截止
            for (int i = 0; i < tableHeadRowEndRowIndexList.Count; i++)
            {
                //列表开始行
                int tableStartIndex = tableHeadRowEndRowIndexList[i];
                while (true)
                {
                    tableStartIndex--;
                    var startRow = table.Rows[tableStartIndex];

                    if (tableStartIndex == 0)
                    {
                        tableStartRowIndexList.Add(tableStartIndex);
                        break;
                    }

                    if (string.IsNullOrWhiteSpace(startRow.RowContent.Trim()) || startRow.RowContent.IsWordTableDateRow())
                    {
                        //取空行或者数据行的下一行当表头开始行
                        tableStartIndex++;
                        tableStartRowIndexList.Add(tableStartIndex);
                        break;
                    }
                }
            }

            //组装分割table
            for (int i = 0; i < tableStartRowIndexList.Count; i++)
            {
                var tableStartIndex = tableStartRowIndexList[i];
                var tableHeadEndIndex = tableHeadRowEndRowIndexList[i];
                //取下一个表格的开始表头的上一行当当前表格的结束行
                var tableEndIndex = i < tableStartRowIndexList.Count - 1 ? tableStartRowIndexList[i + 1] - 1 : table.Rows.Count - 1;
                var newTable = new WordTable()
                {
                    PageNumber = table.PageNumber,
                    IsMatchWordParagraph = table.IsMatchWordParagraph,
                    OperationType = table.OperationType,
                    TableSourceType = TableSourceTypeEnum.SplitTable
                };
                for (int r = tableStartIndex; r <= tableEndIndex; r++)
                {
                    var row = table.Rows[r];
                    //如果包含标题 跳过
                    if (!string.IsNullOrWhiteSpace(row.RowContent.MatchWordTitle()))
                    {
                        continue;
                    }
                    row.RowCells.ForEach(f =>
                    {

                        f.IsHeadColumn = r <= tableHeadEndIndex;
                    });

                    newTable.Rows.Add(row);
                }

                var newTableFirstRowParagraph = table.ContentParagraphs.FirstOrDefault(w => w.Text.Contains(newTable.FirstRowContent));
                var newTableLastRowParagraph = table.ContentParagraphs.FirstOrDefault(w => w.Text.Contains(newTable.LastRowContent));
                newTable.ContentParagraphs = newTableFirstRowParagraph != null && newTableLastRowParagraph != null ?
                    table.ContentParagraphs.Where(w => w.ParagraphNumber >= newTableFirstRowParagraph.ParagraphNumber && w.PageNumber <= newTableLastRowParagraph.ParagraphNumber).ToList()
                    : new List<WordParagraph>();
                if (!newTable.ContentParagraphs.Any())
                {
                    newTable.ContentParagraphs = table.ContentParagraphs;
                }
                tableList.Add(newTable);
            }

            return tableList;
        }

        private static bool IsParaUnderline(WordParagraph paragraph)
        {
            var txt = paragraph.OldText.TrimEnd('\r');
            if (txt.Contains("\t") && txt.Contains("___"))
            {
                return true;
            }
            return false;
        }

        private static bool isCurrencyRow(string[] cells)
        {
            if (cells.Length <= 1)
            {
                return false;
            }
            else
            {
                //正常情况下货币单位在最后一个单元格，但也有一些表格最后是%，向前推
                for (int i = cells.Length - 1; i > 0; i--)
                {
                    string cell = cells[i];
                    if (cell.Trim().Equals("%"))
                        continue;
                    else
                        return isCurrencyCell(cell);

                }
                return false;
            }
        }

        public static bool isCurrencyCell(string cell)
        {
            if (cell == null)
            {
                return false;
            }
            else
            {
                if (cell.Contains("人民币元") || cell.Contains("人民币"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public static bool isBlank(WordParagraph p)
        {
            if (string.IsNullOrEmpty(p.OldText.Trim()))
                return true;
            else
                return false;
        }

        private static int getTabStopsCount(WordParagraph paragraph, bool ignoreInvalid = false)
        {
            return Regex.Matches(paragraph.OldText, "\t").Count;
            //if (paragraph.TabStops == null)
            //{
            //    return -1;
            //}
            //var result = paragraph.TabStops.Count;
            //if (ignoreInvalid)
            //{
            //    float indent = getIndent(paragraph);
            //    TabStops tabstops = paragraph.TabStops;
            //    var tabstopsCount = tabstops.Count;
            //    for (int i = 1; i <= tabstopsCount; i++)
            //    {
            //        // 删除缩进前面的
            //        if (tabstops[i].Position < indent)
            //        {
            //            result--;
            //        }
            //    }
            //}
            //return result;
        }

        private static float getIndent(Microsoft.Office.Interop.Word.Paragraph paragraph)
        {
            /**
         * 悬挂缩进：实际缩进值为 IndentLeft - hanging 首行缩进与左缩进可以同时存在，实际缩进值为IndentLeft +
         * firstLineIndent
         */
            int hanging = paragraph.HangingPunctuation;
            float firstLineIndent = paragraph.FirstLineIndent;
            float IndentLeft = paragraph.LeftIndent;
            hanging = hanging == -1 ? 0 : hanging;
            firstLineIndent = firstLineIndent == -1 ? 0 : firstLineIndent;
            IndentLeft = IndentLeft == -1 ? 0 : IndentLeft;
            if (hanging != 0)
            {
                firstLineIndent = IndentLeft - hanging;
            }
            else
            {
                firstLineIndent = IndentLeft + firstLineIndent;
            }

            return firstLineIndent;
        }

        #endregion





    }
}


