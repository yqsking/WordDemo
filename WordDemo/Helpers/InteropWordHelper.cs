using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using WordDemo;
using Cell = Microsoft.Office.Interop.Word.Cell;


public class InteropWordHelper
{

    public static void ReplaceTextByXml(string wordPath)
    {
        Application wordApp = new Application();
        Document doc = wordApp.Documents.Open(wordPath, ReadOnly: false, Visible: false);
        doc.Activate();
        try
        {
            foreach(Microsoft.Office.Interop.Word.Table table in doc.Tables)
            {
                int rowCount = table.Rows.Count;
                int columnCount = table.Columns.Count;
                for(int rowIndex=1; rowIndex <= rowCount; rowIndex++)
                {
                    for(int columnIndex=1; columnIndex <= columnCount; columnIndex++)
                    {
                        try
                        {
                            Cell cell = table.Cell(rowIndex, columnIndex);
                            Console.ForegroundColor = ConsoleColor.Blue;
                            string cellValue= cell.Range.Text.Replace("\r", "").Replace("\a", "");
                            Console.WriteLine($"{cellValue}({rowIndex},{columnIndex}):宽{cell.Width},高:{cell.Height}");
                            Console.ResetColor();
                            string cellXml = cell.Range.WordOpenXML;
                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.LoadXml(cell.Range.WordOpenXML);
                        }
                        catch(Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine();
                            Console.WriteLine($"({rowIndex},{columnIndex})"+ ex.Message);
                            Console.ResetColor();
                        }
                       
                    }
                }
            }
        }
        catch(Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine("执行异常："+JsonConvert.SerializeObject(ex));
            Console.ResetColor();
        }
        finally
        {
            doc.Save();
            doc.Close();
            wordApp.Quit();
        }
    }

    public static void ReplaceText(string wordPath)
    {
        Application wordApp = new Application();
        Document doc = wordApp.Documents.Open(wordPath, ReadOnly: false, Visible: false);
        doc.Activate();
        try
        {
            int index = 1;
            // 遍历文档中的所有表格
            foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
            {
                string tableXml= table.Range.WordOpenXML;
                var rowDic = new Dictionary<int, List<WordDemo.WordTableCell>>();
                int rowCount = table.Rows.Count;
                int columnCount = table.Columns.Count;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"第{index}个表格:{rowCount}行,{columnCount}列");
                Console.ResetColor();

                //读取所有单元格
                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    var cellValues = new List<string>();
                    var rowCells = new List<WordDemo.WordTableCell>();
                    for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                    {
                        try
                        {
                            Cell cell = table.Cell(rowIndex, columnIndex);
                            cell.Select();
                            var s= doc.Application.Selection.Information[WdInformation.wdStartOfRangeRowNumber];
                            var e = doc.Application.Selection.Information[WdInformation.wdEndOfRangeRowNumber];
                            Console.WriteLine(cell.Height);
                            cellValues.Add(cell.Range.Text.Replace("\r", "").Replace("\a", "")+$"(RowIndex:{cell.RowIndex},ColumnIndex:{cell.ColumnIndex},Width:{cell.Width})");
                            var replaceCell = new WordDemo.WordTableCell()
                            {
                                OldValue = cell.Range.Text.Replace("\r", "").Replace("\a", ""),
                                StartRowIndex = cell.RowIndex,
                                StartColumnIndex = cell.ColumnIndex,
                            };
                            rowCells.Add(replaceCell);
                        }
                        catch
                        {
                           
                        }


                    }

                    rowDic.Add(rowIndex, rowCells);

                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine(string.Join(" ", cellValues));
                    Console.WriteLine();
                    Console.ResetColor();
                }
                ++index;

             
                foreach(var row in rowDic)
                {
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"第{row.Key}行:");
                    Console.ForegroundColor = ConsoleColor.Red;
                    foreach(var cell in row.Value)
                    {
                        //先判断当前行列数是否已经符合最大列数
                        var currentRowColumnCount= row.Value.Count + row.Value.Sum(s => s.ColSpan);
                        if(currentRowColumnCount < columnCount)
                        {
                            int mergeCount = GetHorizontalMergeCount(rowDic, cell.StartRowIndex, cell.StartColumnIndex, columnCount);
                            cell.ColSpan = mergeCount;
                        }

                        Console.WriteLine($"第{cell.StartColumnIndex}个单元格({cell.OldValue})水平合并了个{cell.ColSpan}单元格");
                    }
                }
                Console.WriteLine();

            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine("执行异常：" + JsonConvert.SerializeObject(ex));
            Console.ResetColor();
        }
        finally
        {
            doc.Save();
            doc.Close();
            wordApp.Quit();
        }


    }

    #region

    public void ReplaceText(string filePath, ValueTuple<string[], List<ValueTuple<string, string>>> replaceTuple)
    {
        Application wordApp = new Application();
        Document doc = wordApp.Documents.Open(filePath, ReadOnly: false, Visible: false);
        doc.Activate();
        int index = 1;
        foreach (Paragraph paragraph in doc.Paragraphs)
        {

            ++index;

            if (replaceTuple.Item1.All(w => paragraph.Range.Text.Contains(w)))
            {
                Console.WriteLine($"替换第{index}段:" + paragraph.Range.Text);
                foreach (var item in replaceTuple.Item2)
                {
                    object key = item.Item1;
                    object value = item.Item2;
                    paragraph.Range.Select();
                    FindAndReplace(paragraph.Range, key, value);
                }

            }
        }
        doc.Save();
        doc.Close();
        wordApp.Quit();
    }

    #endregion



    /// <summary>
    /// 获取水平合并单元格数量
    /// </summary>
    /// <param name="rowDic"></param>
    /// <param name="rowIndex"></param>
    /// <param name="columnIndex"></param>
    /// <param name="tableColumnCount"></param>
    /// <returns></returns>
    private static int GetHorizontalMergeCount(Dictionary<int,List<WordDemo.WordTableCell>> rowDic,int rowIndex,int columnIndex,int tableColumnCount)
    {
        int mergeCount = 0;

        //var currentCell= rowDic[rowIndex].FirstOrDefault(w=>w.StartColumnIndex==columnIndex);
        
        ////获取符合表格最大列的行对应单元格
        //var fullColumnRows= rowDic.Where(w => w.Value.Count == tableColumnCount).ToList();
        //foreach(var row in fullColumnRows)
        //{
        //    var cell= row.Value.FirstOrDefault(w=>w.StartColumnIndex==columnIndex);
        //    if(cell==null)
        //    {
        //        continue;
        //    }
            
        //    if(currentCell.CellBorderWidth==cell.CellBorderWidth)
        //    {
        //        //没有水平合并单元格的行 同列的边界宽度与当前单元格的边界宽度一致 说明当前单元格没有合并
        //        break;
        //    }
        //    //取右边的单元格 循环判断边界宽度是否与当前单元格边界宽度一致
        //    var rightCells= row.Value.Where(w => w.StartColumnIndex > cell.StartColumnIndex).OrderBy(o => o.StartColumnIndex).ToList();
        //    foreach(var rightCell in rightCells)
        //    {
        //        if(currentCell.CellBorderWidth==rightCell.CellBorderWidth)
        //        {
        //            mergeCount = rightCell.StartColumnIndex - cell.StartColumnIndex + 1;
        //            break;
        //        }
        //    }

        //}

        return mergeCount;
    }

    /// <summary>
    /// 获取垂直合并单元格数量
    /// </summary>
    /// <param name="table"></param>
    /// <param name="startRow"></param>
    /// <param name="col"></param>
    /// <param name="content"></param>
    /// <returns></returns>
    private static int GetVerticalMergeCount(Microsoft.Office.Interop.Word.Table table, int startRow, int col, string content)
    {
        int mergeCount = 0;

        int count = 1;

        // 从下一行开始检查，直到内容改变或到达表格底部
        for (int r = startRow + 1; r <= table.Rows.Count; r++)
        {
            // 首先，安全地尝试获取下一个单元格
            Cell nextCell = null;
            try
            {
                nextCell = table.Cell(r, col);
            }
            catch
            {
                // 如果访问异常，说明可能遇到了合并边界，直接跳出循环
                break;
            }

            // 比较内容，注意这里需要确保nextCell不为空
            if (nextCell != null && nextCell.Range.Text.Trim() == content)
            {
                count++;
            }
            else
            {
                break;
            }
        }

        mergeCount = count - 1;
        return mergeCount;
    }

    /// <summary>
    /// 替换范围文本
    /// </summary>
    /// <param name="range"></param>
    /// <param name="findText"></param>
    /// <param name="replaceWithText"></param>
    private void FindAndReplace(Range range, object findText, object replaceWithText)
    {

        object matchCase = true;
        object matchWholeWord = false;
        object matchWildCards = true;
        object matchSoundsLike = false;
        object matchAllWordForms = false;
        object forward = true;
        object format = false;
        object matchKashida = false;
        object matchDiacritics = false;
        object matchAlefHamza = false;
        object matchControl = false;
        object replace = 2; //替换规则 0：不替换 1：替换第一个 2：替换所有
        object wrap = 0;

        range.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
        ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText,
        ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
    }



}
