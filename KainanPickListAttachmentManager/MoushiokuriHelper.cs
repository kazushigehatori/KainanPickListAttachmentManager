using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;



namespace KainanPickListAttachmentManager.Helpers
{
    internal class MoushiokuriHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private static AdvanceSoftware.ExcelCreator.Creator creator1 = new AdvanceSoftware.ExcelCreator.Creator();
        private static AdvanceSoftware.ExcelCreator.Creator creator21 = new AdvanceSoftware.ExcelCreator.Creator();

        /// <summary>
        /// デフォルトの申送りシート名
        /// </summary>
        private const string DefaultMoushiokuriSheet = "申送り";

        /// <summary>
        /// 診療科名から対応する申送りシート名を取得する（動的検索版）
        /// テンプレート内のシート名を動的に検索して対応するシートを特定する
        /// EmpowerORから渡される診療科名は先頭5文字に切り詰められているため、
        /// 先頭5文字での部分一致で判定する
        /// </summary>
        /// <param name="creator">ExcelCreatorオブジェクト（テンプレートが開かれている状態）</param>
        /// <param name="clinSecNm">診療科名（最大5文字に切り詰められている場合あり）</param>
        /// <returns>対応する申送りシート名</returns>
        private static string GetMoushiokuriSheetName(AdvanceSoftware.ExcelCreator.Creator creator, string clinSecNm)
        {
            string inputDept = clinSecNm?.Trim() ?? "";
            if (string.IsNullOrEmpty(inputDept))
            {
                logger.Info($"診療科名が空のため、デフォルトシート '{DefaultMoushiokuriSheet}' を使用します。");
                return DefaultMoushiokuriSheet;
            }

            // 入力が5文字を超える場合は先頭5文字を使用
            string inputKey = inputDept.Length > 5 ? inputDept.Substring(0, 5) : inputDept;
            logger.Info($"診療科名検索: 入力='{clinSecNm}', 検索キー='{inputKey}'");

            // テンプレート内の全シートをループ
            int sheetCount = creator.SheetCount;
            for (int i = 0; i < sheetCount; i++)
            {
                creator.SheetNo = i;
                string sheetName = creator.SheetName;

                // 「申送り」で始まり、かつ「申送り」そのものではないシート
                if (sheetName.StartsWith("申送り") && sheetName != DefaultMoushiokuriSheet)
                {
                    // 「申送り」以降の部分を取得（診療科名部分）
                    string sheetDept = sheetName.Substring(3); // "申送り" は3文字

                    // 診療科名部分の先頭5文字と比較
                    string sheetKey = sheetDept.Length > 5 ? sheetDept.Substring(0, 5) : sheetDept;

                    if (inputKey == sheetKey)
                    {
                        logger.Info($"診療科 '{clinSecNm}' (検索キー: '{inputKey}') に対応するシート '{sheetName}' を選択しました。");
                        return sheetName;
                    }
                }
            }

            logger.Info($"診療科 '{clinSecNm}' に対応するシートが見つからないため、デフォルトシート '{DefaultMoushiokuriSheet}' を使用します。");
            return DefaultMoushiokuriSheet;
        }

        /// <summary>
        /// 指定されたシート以外の「申送り」で始まるシートを削除する
        /// </summary>
        /// <param name="creator">ExcelCreatorオブジェクト</param>
        /// <param name="activeSheetName">残すシート名</param>
        private static void DeleteMoushiokuriSheetsExcept(AdvanceSoftware.ExcelCreator.Creator creator, string activeSheetName)
        {
            int sheetCount = creator.SheetCount;
            logger.Info($"シート数: {sheetCount}, 残すシート: '{activeSheetName}'");

            // 削除対象のシート名リストを作成（削除時にインデックスがずれるため、先にリスト化）
            List<string> sheetsToDelete = new List<string>();

            for (int i = 0; i < sheetCount; i++)
            {
                creator.SheetNo = i;
                string sheetName = creator.SheetName;

                if (sheetName.StartsWith("申送り") && sheetName != activeSheetName)
                {
                    sheetsToDelete.Add(sheetName);
                    logger.Info($"削除対象シート: '{sheetName}'");
                }
            }

            // シート名から番号を取得して削除（毎回最新の番号を取得）
            foreach (string sheetName in sheetsToDelete)
            {
                int sheetNo = creator.SheetNo2(sheetName);
                if (sheetNo >= 0)
                {
                    creator.DeleteSheet(sheetNo, 1);
                    logger.Info($"シート '{sheetName}' (No.{sheetNo}) を削除しました。");
                }
                else
                {
                    logger.Warn($"シート '{sheetName}' が見つかりませんでした。");
                }
            }

            logger.Info($"削除後のシート数: {creator.SheetCount}");
        }



        private static void creator1_Error(object sender,
            AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator1: Error processing file");
            logger.Error("creator1: " + creator1.ErrorMessage);
        }

        private static void creator21_Error(object sender,
    AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator21: Error processing file");
            logger.Error("creator21: " + creator21.ErrorMessage);
        }


        public static bool InsertMoushiokuriItems(List<string> filePaths)
        {
            string OpeDate_cell = System.Configuration.ConfigurationManager.AppSettings["Moushiokuri_OpeDate_cell"];
            string PatientID_cell = System.Configuration.ConfigurationManager.AppSettings["Moushiokuri_PatientID_cell"];
            string PatientName_cell = System.Configuration.ConfigurationManager.AppSettings["Moushiokuri_PatientName_cell"];
            string ClinicalDept_cell = System.Configuration.ConfigurationManager.AppSettings["Moushiokuri_ClinicalDept_cell"];
            string OpeMethodSetCode_cell = System.Configuration.ConfigurationManager.AppSettings["Moushiokuri_OpeMethodSetCode_cell"];
            string OpeMethodSetName_cell = System.Configuration.ConfigurationManager.AppSettings["Moushiokuri_OpeMethodSetName_cell"];

            System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense = myAssembly.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator1.SetLicense(tempStreamLicense);

            creator1.Error += new AdvanceSoftware.ExcelCreator.CreatorErrorEventHandler(creator1_Error);

            System.Reflection.Assembly myAssembly21 = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense21 = myAssembly21.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator21.SetLicense(tempStreamLicense21);

            creator21.Error += new AdvanceSoftware.ExcelCreator.CreatorErrorEventHandler(creator21_Error);



            foreach (var filePath in filePaths)
            {
                try
                {
                    //
                    // 手術準備表は filePathに渡ってくるファイル名を"申残"を"ＰＬ"に変換したもの
                    //
                    string filePath2 = filePath.Replace("申残", "ＰＬ");

                    //
                    // ExcelCreatorオブジェクトの作成
                    // ＰＬファイルを開く
                    //
                    creator1.OpenBook(filePath2, "");
                    //
                    // ExcelCreatorオブジェクト21の作成
                    // 申残ファイルを開く
                    //
                    creator21.OpenBook(filePath, "");

                    // シート"1"を選択
                    creator1.SheetNo = 0;

                    // 名前付きセルから値を取得
                    string OpeDateValue = (string)creator1.Cell("YtiOpeDate").Value;
                    string PatientIDValue = (string)creator1.Cell("PatientID").Value;
                    string PatientNameValue = (string)creator1.Cell("PatientNm").Value;
                    string ClinicalDeptValue = (string)creator1.Cell("ClinSecNm").Value;
                    string OpeMethodSetCodeValue = (string)creator1.Cell("OpeMethSetCD").Value;
                    string OpeMethodSetNameValue = (string)creator1.Cell("OpeMethSetNm").Value;


                    // 診療科名から対応する申送りシート名を取得（テンプレートから動的に検索）
                    string targetSheetName = GetMoushiokuriSheetName(creator21, ClinicalDeptValue);

                    // 該当シート以外の「申送り」で始まるシートを削除
                    DeleteMoushiokuriSheetsExcept(creator21, targetSheetName);

                    // 対応する申送りシートを選択
                    int MoushiokuriNo = creator21.SheetNo2(targetSheetName);
                    if (MoushiokuriNo < 0)
                    {
                        // シートが見つからない場合はデフォルトにフォールバック
                        logger.Warn($"シート '{targetSheetName}' が見つかりません。デフォルトシート '{DefaultMoushiokuriSheet}' を使用します。");
                        MoushiokuriNo = creator21.SheetNo2(DefaultMoushiokuriSheet);
                        // デフォルトシート以外を削除
                        DeleteMoushiokuriSheetsExcept(creator21, DefaultMoushiokuriSheet);
                    }
                    creator21.SheetNo = MoushiokuriNo;

                    //
                    // OpeDateValueに 西暦が入っていたら削る
                    //
                    int indexOfYear = OpeDateValue.IndexOf("年");
                    if (indexOfYear != -1)
                    {
                        // "年"の位置より前を削除（"年"以降を取得）
                        OpeDateValue = OpeDateValue.Substring(indexOfYear + 1) + " ";
                    }
                    OpeDateValue = OpeDateValue.Replace("月", " 月 ");
                    OpeDateValue = OpeDateValue.Replace("日", " 日 ");


                    // 指定セルに値を設定
                    creator21.Cell(OpeDate_cell).Str = OpeDateValue;
                    creator21.Cell(PatientID_cell).Str = PatientIDValue;
                    creator21.Cell(PatientName_cell).Str = PatientNameValue;
                    creator21.Cell(ClinicalDept_cell).Str = ClinicalDeptValue;
                    creator21.Cell(OpeMethodSetCode_cell).Str = OpeMethodSetCodeValue;
                    creator21.Cell(OpeMethodSetName_cell).Str = OpeMethodSetNameValue;

                    // 保存して閉じる
                    creator1.CloseBook(false);
                    creator21.CloseBook(true);
                }
                catch (Exception ex)
                {
                    // ログ出力など必要に応じて
                    Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
                    return false;
                }

            }
            return true;
        }
    }
}
