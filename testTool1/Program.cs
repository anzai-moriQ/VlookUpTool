using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;

// エクセルの表示・非表示、また起動・停止を司る
namespace testTool1 
{
    class Program
    {


        static void Main(string[] args)
        {
            ValueGetter valueGetter = new ValueGetter();    // Excelブック内のデータを取得するインスタンス
            Collating   collating   = new Collating();      // 突合処理を行うインスタンス

            
            // テスト用
            //string filePath = @"C:\Users\teramoto.yuki\Documents\コンソールアプリのドキュメント\標的.xlsx";
            //string targetSheetName = "検索対象";
            //string outSheetName = "出力結果";


            //// 必須入力項目
            Console.WriteLine("処理対象のエクセルファイルの絶対パスを入力してください");
            string filePath = Console.ReadLine();

            // ファイルパスからディレクトリを取得
            string directory = Path.GetDirectoryName(filePath);

            //// 必須入力項目
            Console.WriteLine("検索対象シート名を入力してください");
            string targetSheetName = Console.ReadLine();

            //// 必須入力項目
            Console.WriteLine("結果を出力するためのシートを指定してください");
            string outSheetName = Console.ReadLine();

            try
            {
                // 必須入力項目チェック
                if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(targetSheetName) || string.IsNullOrWhiteSpace(outSheetName))
                {
                    throw new Exception("入力されていない項目がありました");
                }

                // 例外を未然に防ぐためのディレクトリチェック
                if (Directory.Exists(directory))
                {
                    //Excelブック・シートの取得
                    XLWorkbook workBook = new XLWorkbook(filePath);
                    IXLWorksheet workSheet = workBook.Worksheet(1);                     // 検索条件となるワークシート
                    IXLWorksheet tgWorkSheet = workBook.Worksheet(targetSheetName);     // 検索対象となるワークシート
                    IXLWorksheet outputSheet;

                    if (workBook.TryGetWorksheet(outSheetName, out outputSheet) == false)
                    {
                        // 出力用シートがなければ二番目に作成する
                        workBook.AddWorksheet(outSheetName, 2);
                    }
                    outputSheet = workBook.Worksheet(outSheetName);                     // 結果出力用のワークシート


                    // メンバ変数
                    var criteriaList = new Dictionary<int, List<string>>();
                    var targetList = new Dictionary<int, List<string>>();
                    var collatingList = new List<List<string>>();


                    /***
                     * ここからデータ取得処理
                     * ValueGetterクラス呼び出し
                     * 戻り値:Dictionary<int, List<string>>型
                     ***/
                    criteriaList = valueGetter.ValueGet(workSheet);     // 検索条件となるデータを取得する

                    // 検索条件となるデータが0件の場合ここで処理を終了する
                    if (criteriaList.Count < 1)
                    {
                        Console.WriteLine("検索条件となるデータが0件でした");
                        return;
                    }

                    targetList = valueGetter.ValueGet(tgWorkSheet);     // 検索対象のデータを取得

                    // 検索対象となるデータが0件の場合ここで処理を終了する
                    if (targetList.Count < 1)
                    {
                        Console.WriteLine("検索対象となるデータが0件でした");
                        return;
                    }


                    /***
                     * ここから検索処理(突合)
                     * collatingクラス呼び出し
                     * 戻り値:List<List<string>>型
                     ***/
                    collatingList = collating.Collatings(criteriaList, targetList);


                    /***
                     * ここからoutputSheetへのデータ吐き出し処理
                     ***/
                    int c = 1;
                    foreach (var item in collatingList)
                    {
                        int t = 1;
                        foreach (var child in item)
                        {
                            outputSheet.Cell(c, t).Value = child;
                            t++;
                        }
                        c++;
                    }


                    workBook.Save();

                    Console.WriteLine("処理終了のお知らせ");
                }
                else
                {
                    throw new Exception("存在しないディレクトリが指定されました");
                }
            }
            catch (NameNotRecognizedException)
            {
                Console.WriteLine("NameNotRecognizedException：");
                Console.WriteLine("対応していない関数がExcelファイルに含まれています");
            }
            catch (ArgumentException)
            {
                Console.WriteLine("ArgumentException：");
                Console.WriteLine("以下の原因が考えられます");
                Console.WriteLine("　　　-- ファイルが存在しません");
                Console.WriteLine("　　　-- 拡張子が不正です");
                Console.WriteLine("　　　-- 検索対象のシートが存在しません");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

    }
}