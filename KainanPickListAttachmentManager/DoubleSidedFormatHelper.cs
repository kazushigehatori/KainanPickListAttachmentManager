using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AdvanceSoftware.ExcelCreator;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace KainanPickListAttachmentManager.Helpers
{
    public class DoubleSidedFormatHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static AdvanceSoftware.ExcelCreator.Creator creator33 = new AdvanceSoftware.ExcelCreator.Creator();

        private static void creator33_Error(object sender,
            AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator33: Error processing file");
            logger.Error("creator33: " + creator33.ErrorMessage);
        }


        public static bool InsertBlankPages(List<string> filePaths)
        {
            System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense = myAssembly.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator33.SetLicense(tempStreamLicense);

            foreach (var path in filePaths)
            {
                string filepath = path.Replace("申残","ＰＬ");

                //
                // 5回Maxにする
                //
                creator33.OpenBook(filepath, filepath);
                for (int count = 0; count < 5; count++)
                {
                    //
                    // シート数を得る
                    //
                    int SheetCnt = creator33.SheetCount;
                    for (int i = 0; i < SheetCnt; i++)
                    {
                        //
                        // シートNoを設定
                        //
                        creator33.SheetNo = i;
                        if ((creator33.Pos(0, 0).Str).Contains("手術") && (i % 2 == 1))
                        {
                            creator33.AddSheet(1, i);
                            creator33.SheetNo = i;
                            creator33.Pos(0, 0).Str = " ";
                            break;
                        }
                    }
                    
                }
                creator33.CloseBook(true);
            }



            return true;
        }

    }


}
