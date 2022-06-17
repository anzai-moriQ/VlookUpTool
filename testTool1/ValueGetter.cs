using System;
using System.IO;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testTool1
{
    /*
     * Excelファイル内のデータを取得するクラス
     */
    internal class ValueGetter
    {
        public Dictionary<int, List<string>> ValueGet(IXLWorksheet workSheet)
        {

            var dic = new Dictionary<int, List<string>>();
            if (workSheet.LastRowUsed() != null)
            {
                int lastRow = workSheet.LastRowUsed().RowNumber();
                int lastColumn = workSheet.LastColumnUsed().ColumnNumber();
                for (int i = 1; i <= lastRow; i++)
                {
                    var list = new List<string>();
                    for (int j = 1; j <= lastColumn; j++)
                    {
                        IXLCell cell = workSheet.Cell(i, j);

                        // 数値が想定されているが、扱いやすさ向上のため一旦stringにキャストする
                        list.Add(cell.Value.ToString());
                    }
                    dic.Add(i, list);
                }
            }
            return dic;
        }
    }
}
