using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testTool1
{
    /*
     * シート間のデータを照合するクラス
     * シート１のA列をキーとする
     */
    internal class Collating
    {
        public List<List<string>> Collatings(Dictionary<int, List<string>> criteriaList, Dictionary<int, List<string>> targetList)
        {
            var collList = new List<List<string>>();
            for (int i = 1; i <= criteriaList.Count; i++)
            {
                // 検索条件となるValue値
                var item = criteriaList[i][0];

                for (int j = 1; j <= targetList.Count; j++)
                {
                    /*
                     * Dictionaryの中に格納されているValueを一つずつ確認する
                     * Valueの項目数は可変のためループで確認する必要有
                     */
                    for (int k = 0; k < targetList[i].Count; k++)
                    {
                        // 検索対象となるValue値
                        var target = targetList[j][k];
                        if (item == target)
                        {
                            collList.Add(targetList[i]);
                            break;
                        }
                    }
                }
            }

            return collList;
        }
    }
}
