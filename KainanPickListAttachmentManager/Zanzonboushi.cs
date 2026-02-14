using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AdvanceSoftware.ExcelCreator;
using Microsoft.Office.Interop.Excel;
using NLog;




namespace KainanPickListAttachmentManager.Helpers
{
    internal class ZanzonboushiHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private static AdvanceSoftware.ExcelCreator.Creator creator2 = new AdvanceSoftware.ExcelCreator.Creator();
        private static AdvanceSoftware.ExcelCreator.Creator creator22 = new AdvanceSoftware.ExcelCreator.Creator();



        private static void creator2_Error(object sender,
            AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator2: Error processing file");
            logger.Error("creator2: " + creator2.ErrorMessage);
        }

        private static void creator22_Error(object sender,
             AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator22: Error processing file");
            logger.Error("creator22: " + creator22.ErrorMessage);
        }


        public static bool InsertZanzonboushiItems(List<string> filePaths)
        {

            //
            // ＰＬ用
            //
            System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense = myAssembly.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator2.SetLicense(tempStreamLicense);

            creator2.Error += new AdvanceSoftware.ExcelCreator.CreatorErrorEventHandler(creator2_Error);

            //
            // 申残用
            //
            System.Reflection.Assembly myAssembly22 = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense22 = myAssembly22.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator22.SetLicense(tempStreamLicense22);

            creator22.Error += new AdvanceSoftware.ExcelCreator.CreatorErrorEventHandler(creator22_Error);


            foreach (var filePath in filePaths)
            {
                try
                {
                    //
                    // 手術準備表は filePathに渡ってくるファイル名を"申残"を"ＰＬ"に変換したもの
                    //
                    string filePath2 = filePath.Replace("申残", "ＰＬ");

                    // ExcelCreatorオブジェクトの作成ＰＬ
                    creator2.OpenBook(filePath2, "");

                    // ExcelCreatorオブジェクトの作成申残
                    creator22.OpenBook(filePath, "");

                    //
                    // 残存防止シート、ヘッダ部分の挿入を行う
                    //
                    InsertHeaderItems();

                    //
                    // 残存防止、残存防止商品リストをピッキングリストより取得
                    //
                    List<string> ZanzonBoushiGoodsList = new List<string>();

                    //
                    // シート数の取得 creator2 xlsx
                    //
                    int SheetCnt = creator2.SheetCount;

                    //
                    // シート数分回す
                    //
                    for (int i = 0; i < SheetCnt; i++)
                    {
                        //
                        // シートNoを設定
                        //
                        creator2.SheetNo = i;
                        
                        //
                        // 手術の文字が(0,0)に記載されていれば1P目とみなす
                        //
                        if ( (creator2.Pos(0, 0).Str).Contains("手術") )
                        {
                            GetP1_ZanzonBoushiGoods(ZanzonBoushiGoodsList);
                        }
                        else
                        {
                            GetP2Onward_ZanzonBoushiGoods(ZanzonBoushiGoodsList);
                        }
                    }


                    //
                    // 残存防止、残存防止シートに設定
                    //
                    SetZanzonBoushiGoodsToSheet(ZanzonBoushiGoodsList);

                    // 保存して閉じる
                    creator2.CloseBook(false);
                    creator22.CloseBook(true);
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


        //
        // 残存防止シート、ヘッダ部分の挿入を行う
        //
        private static bool InsertHeaderItems()
        {
            string OpeDate_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_OpeDate_cell"];
            string PatientID_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PatientID_cell"];
            string PatientName_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PatientName_cell"];
            string ClinicalDept_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_ClinicalDept_cell"];
            string OpeMethodSetCode_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_OpeMethodSetCode_cell"];
            string OpeMethodSetName_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_OpeMethodSetName_cell"];

            //
            // Excel手術準備表ＰＬについて
            // シート"1"を選択
            //
            creator2.SheetNo = 0;

            // 名前付きセルから値を取得
            string OpeDateValue = (string)creator2.Cell("YtiOpeDate").Value;
            string PatientIDValue = (string)creator2.Cell("PatientID").Value;
            string PatientNameValue = (string)creator2.Cell("PatientNm").Value;
            string ClinicalDeptValue = (string)creator2.Cell("ClinSecNm").Value;
            string OpeMethodSetCodeValue = (string)creator2.Cell("OpeMethSetCD").Value;
            string OpeMethodSetNameValue = (string)creator2.Cell("OpeMethSetNm").Value;


            // シート"残存防止"を選択
            int ZanzonboushiNo = creator22.SheetNo2("残存防止");
            creator22.SheetNo = ZanzonboushiNo;


            //
            // OpeDateValueに 西暦が入っていたら削る
            //
            int indexOfYear = OpeDateValue.IndexOf("年");
            if (indexOfYear != -1)
            {
                // "年"の位置より前を削除（"年"以降を取得）
                OpeDateValue = OpeDateValue.Substring(indexOfYear + 1);

            }
            OpeDateValue = OpeDateValue.Replace("月", " 月 ");
            OpeDateValue = OpeDateValue.Replace("日", " 日 ");

            // 指定セルに値を設定
            creator22.Cell(OpeDate_cell).Str = OpeDateValue;
            creator22.Cell(PatientID_cell).Str = PatientIDValue;
            creator22.Cell(PatientName_cell).Str = PatientNameValue;
            creator22.Cell(ClinicalDept_cell).Str = ClinicalDeptValue;
            creator22.Cell(OpeMethodSetCode_cell).Str = OpeMethodSetCodeValue;
            creator22.Cell(OpeMethodSetName_cell).Str = OpeMethodSetNameValue;
            return true;
        }


        //
        // ページ1
        //
        private static bool GetP1_ZanzonBoushiGoods(List<string> GoodsList)
        {
            int LineCount = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL1P_Count"]);
            int nStep = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL1P_Step"]);
            string GoodsNameStart_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL1P_GoodsNameStart_cell"];
            string GoodsCnt2Start_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL1P_GoodsCnt2Start_cell"];
            string GoodsCntStart_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL1P_GoodsCntStart_cell"];
            string BpnCDStart_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL1P_BpnCDStart_cell"];


            var (colgoods, rowgoods) = ParseCell(GoodsNameStart_cell);
            var (colcnt2, rowcnt2) = ParseCell(GoodsCnt2Start_cell);
            var (colcnt, rowcnt) = ParseCell(GoodsCntStart_cell);
            var (colbpn, rowbpn) = ParseCell(BpnCDStart_cell);

            
            //
            // シート Noは上位で設定済み
            //

            int nrow = rowcnt2;

            for (int i = 0; i < LineCount; i++)
            {
                string cleanedCnt2 = Regex.Replace(creator2.Pos(colcnt2, nrow).Str, @"[\s　]+", "");
                string cleanedCnt = Regex.Replace(creator2.Pos(colcnt, nrow).Str, @"[\s　]+", "");


                int nCnt2 = 0;
                int nCnt = 0;

                if ((int.TryParse(cleanedCnt2, out int num2) && num2 >= 1))
                {
                    nCnt2 = num2;
                }

                if ((int.TryParse(cleanedCnt, out int num) && num >= 1))
                {
                    nCnt = num;
                }

                //
                // GoodsCnt2 == GoodsCntを記載
                // 
                //
                bool isKizai = false;
                string BpnCD = creator2.Pos(colbpn, nrow + (rowbpn - rowcnt2) ).Str;


                //
                // BpnCDにRがあるものはを外す
                //
                if (BpnCD.Length >= 2)
                {
                    if (BpnCD[0] == 'R')
                    {
                        isKizai = true;
                    }
                    else if (BpnCD[1] == 'R')
                    {
                        isKizai = true;
                    }
                }

                if (nCnt > 0 && (nCnt2 == nCnt) && (isKizai == false ) )
                {
                    string goodstr = creator2.Pos(colgoods, nrow).Str;
                    if (!string.IsNullOrEmpty(goodstr))
                    {
                        //
                        // Goodstr末尾が '*'のものは外す
                        //
                        string cleaned = Regex.Replace(goodstr, @"\r\n|\r|\n", " ");
                        if (!cleaned.EndsWith("*"))
                        {
                            GoodsList.Add(cleaned);
                        }
                    }
                }

                //
                // nrowをnStepでインクリメント
                //
                nrow += nStep;
            }

            return true;
        }



        //
        // ページ2以降
        //
        private static bool GetP2Onward_ZanzonBoushiGoods(List<string> GoodsList)
        {
            int LineCount = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL2P_Count"]);
            int nStep = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL2P_Step"]);
            string GoodsNameStart_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL2P_GoodsNameStart_cell"];
            string GoodsCnt2Start_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL2P_GoodsCnt2Start_cell"];
            string GoodsCntStart_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL2P_GoodsCntStart_cell"];
            string BpnCDStart_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_PL2P_BpnCDStart_cell"];

            //
            // シートNoは上位で設定済み
            //



            // 拾い出しの最初を設定
            // 2P以降各ページはすべて同じ
            //
            var (colgoods, rowgoods) = ParseCell(GoodsNameStart_cell);
            var (colcnt2, rowcnt2) = ParseCell(GoodsCnt2Start_cell);
            var (colcnt, rowcnt) = ParseCell(GoodsCntStart_cell);
            var (colbpn, rowbpn) = ParseCell(BpnCDStart_cell);

            //
            //作業拾い出しシート番号の設定
            //
            int nrow = rowcnt2; // 完全未実装
            for (int i = 0; i < LineCount; i++)
            {
                string cleanedCnt2 = Regex.Replace(creator2.Pos(colcnt2, nrow).Str, @"[\s　]+", "");
                string cleanedCnt = Regex.Replace(creator2.Pos(colcnt, nrow).Str, @"[\s　]+", "");


                int nCnt2 = 0;
                int nCnt = 0;

                if ((int.TryParse(cleanedCnt2, out int num2) && num2 >= 1))
                {
                    nCnt2 = num2;
                }

                if ((int.TryParse(cleanedCnt, out int num) && num >= 1))
                {
                    nCnt = num;
                }


                //
                // GoodsCnt2 == GoodsCntを記載

                //


                bool isKizai = false;
                string BpnCD = creator2.Pos(colbpn, nrow + (rowbpn - rowcnt2)).Str;


                //
                // BpnCDにRがあるものはを外す
                //
                if (BpnCD.Length >= 2)
                {
                    if (BpnCD[0] == 'R')
                    {
                        isKizai = true;
                    }
                    else if (BpnCD[1] == 'R')
                    {
                        isKizai = true;
                    }
                }

                if (nCnt > 0 && (nCnt2 == nCnt) && (isKizai == false) )
                {
                    string goodstr = creator2.Pos(colgoods, nrow).Str;
                    if (!string.IsNullOrEmpty(goodstr))
                    {
                        //
                        // Goodstr末尾が '*'のものは外す
                        //
                        string cleaned = Regex.Replace(goodstr, @"\r\n|\r|\n", " ");
                        if (!cleaned.EndsWith("*"))
                        {
                            GoodsList.Add(cleaned);
                        }
                    }
                }


                //
                // nrowをnStepでインクリメント
                //
                nrow += nStep;
            }


            return true;
        }

        private static bool SetZanzonBoushiGoodsToSheet(List<string> GoodsList)
        {
            string Column1Start_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_SheetColumn1Start_cell"];
            int col1Count = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_SheetColumn1_Count"]);

            string Column2Start_cell = System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_SheetColumn2Start_cell"];
            int col2Count = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Zanzonboushi_SheetColumn2_Count"]);


            // シート"残存防止"を選択
            int ZanzonboushiNo = creator22.SheetNo2("残存防止");
            creator22.SheetNo = ZanzonboushiNo;

            //
            // GoodsListを差し込み
            //
            int nth = 0;
            foreach (string goods in GoodsList)
            {
                string destcell = GetZanzonBoushiSheetCellByNum(Column1Start_cell, col1Count, Column2Start_cell, col2Count, nth);
                if (!string.IsNullOrEmpty(destcell))
                {
                    creator22.Cell(destcell).Str = goods;
                }
                //
                // インクリメント
                //
                nth++;
            }
            return true;
        }

        private static string GetZanzonBoushiSheetCellByNum(string C1cell, int C1num, string C2cell, int C2num, int Num)
        {
            if (Num >= C1num + C2num)
            {
                logger.Error("無効なセルNumです。:" + Num);
                return "";
            }
            if (Num < 0)
            {
                logger.Error("無効なセルNumです。:" + Num);
                return "";
            }

            string C1col;
            string C2col;
            int row1;
            int row2;

            if (Num < C1num)
            {
                // 段組み１つめ
                // 正規表現で列と行を分離
                Match match = Regex.Match(C1cell, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    C1col = match.Groups[1].Value; // 列（英文字）
                    row1 = int.Parse(match.Groups[2].Value);    // 行（数字）

                    return C1col + (row1 + Num).ToString();
                }
                else
                {
                    logger.Error("無効なC1セル座標です。:" + C1cell);

                }

            }
            else
            {
                // 段組み２つめ
                // 正規表現で列と行を分離
                Match match = Regex.Match(C2cell, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    C2col = match.Groups[1].Value; // 列（英文字）
                    row2 = int.Parse(match.Groups[2].Value);    // 行（数字）

                    return C2col + (row2 + Num - C1num).ToString();
                }
                else
                {
                    logger.Error("無効なC2セル座標です。:" + C2cell);

                }

            }

            return "";
        }


        public static (int columnIndex, int rowIndex) ParseCell(string cell)
        {
            if (string.IsNullOrEmpty(cell))
            { 
                logger.Error("セル文字列が空です。:" + cell);
                return (0, 0);
            }
            
            Match match = Regex.Match(cell, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                logger.Error("セル文字列の形式が正しくありません。:" + cell);
                return (0, 0);
            }

            string columnLetters = match.Groups[1].Value.ToUpper();
            int rowIndex = int.Parse(match.Groups[2].Value);

            int columnIndex = 0;
            foreach (char c in columnLetters)
            {
                columnIndex = columnIndex * 26 + (c - 'A' + 1);
            }

            columnIndex -= 1; // 0始まりに調整
            rowIndex -= 1; // 0始まりに調整

            return (columnIndex, rowIndex);
        }


    }
}
