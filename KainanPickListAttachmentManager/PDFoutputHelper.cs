using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AdvanceSoftware.ExcelCreator;
using NLog;

namespace KainanPickListAttachmentManager.Helpers
{
    internal class PDFoutputHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private static AdvanceSoftware.ExcelCreator.Creator creator3 = new AdvanceSoftware.ExcelCreator.Creator();

        private static void creator3_Error(object sender,
            AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator3: Error processing file");
            logger.Error("creator2: " + creator3.ErrorMessage);
        }

        public static bool PDFfilesOutput(List<string> filePaths)
        {
            System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense = myAssembly.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator3.SetLicense(tempStreamLicense);

            creator3.Error += new AdvanceSoftware.ExcelCreator.CreatorErrorEventHandler(creator3_Error);



            //
            // 申残PDFの作成
            //
            foreach (var filePath in filePaths)
            {
                // ExcelCreatorオブジェクトの作成
                creator3.OpenBook(filePath, filePath);

                string pdffile = filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                    ? filePath.Substring(0, filePath.Length - 4) + "pdf"
                    : filePath;

                creator3.CloseBook(true, pdffile, false);
            }

            //
            // ＰＬPDFの作成
            //
            foreach (var filePath in filePaths)
            {
                string filePath2 = filePath.Replace("申残", "ＰＬ");
                // ExcelCreatorオブジェクトの作成
                creator3.OpenBook(filePath2, filePath2);

                string pdffile = filePath2.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                    ? filePath2.Substring(0, filePath2.Length - 4) + "pdf"
                    : filePath2;

                creator3.CloseBook(true, pdffile, false);
            }

            return true;
        }

    }
}
