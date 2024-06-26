using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WordDemo
{
    public class OpenXmlWordHelper
    {

        public static void ReplaceText(string wordPath)
        {

            using (WordprocessingDocument document = WordprocessingDocument.Open(wordPath, false))
            {
                MainDocumentPart mainPart = document.MainDocumentPart;
                IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Table> tables = mainPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>();
                foreach (DocumentFormat.OpenXml.Wordprocessing.Table table in tables)
                {
                    int rowIndex = 0;
                    var rows = table.Elements<TableRow>();
                    foreach (TableRow row in rows)
                    {
                        var tableRowDic = new Dictionary<int, List<WordTableCell>>();
                        rowIndex++;
                        $"第{rowIndex}行:".Console(ConsoleColor.Blue);

                        var rowCells = new List<WordTableCell>();
                        int columnIndex = 0;
                        foreach (TableCell cell in row.Elements<TableCell>())
                        {
                            columnIndex++;
                            var paras = cell.Elements<Paragraph>();
                            StringBuilder cellText = new StringBuilder();
                            foreach (var para in paras)
                            {
                                if (para != null && para.InnerText != null)
                                {
                                    cellText.Append(para.InnerText);
                                }
                            }
                            var replaceCell=new WordTableCell() { 
                                OldValue = cellText.ToString(),
                                StartColumnIndex= columnIndex,
                                ColSpan= cell.TableCellProperties.GridSpan?.Val??0,
                                StartRowIndex= rowIndex,
                            };
                           
                            var horizontalMergeValue= cell.TableCellProperties?.HorizontalMerge?.InnerText;
                            var verticalMergeValue= cell.TableCellProperties?.VerticalMerge?.InnerText;
                           // $"{cellText}({rowIndex},{columnIndex})(GridSpan:{gridSpan},HorizontalMergeValue:{horizontalMergeValue},VerticalMergeValue:{verticalMergeValue})".Console(ConsoleColor.Red);
                        }
                       
                    }
                }

            }
        }
    }
}
