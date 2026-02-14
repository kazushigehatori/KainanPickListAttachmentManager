using System;
using System.Collections.Generic;
using System.IO;
using KainanPickListAttachmentManager.Helpers;
using AdvanceSoftware.ExcelCreator;
using NLog;


namespace KainanPickListAttachmentManager
{


    public static class GlobalValues
    {
        public static bool MasterMode = false;
        public static bool MasterOpeRoom = false;
    }

    class Program
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {

            //
            // 引数でMaster Modeを設定する
            //
            if (args.Length > 0)
            {
                if (args[0] == "Master")
                {
                    GlobalValues.MasterMode = true;
                    GlobalValues.MasterOpeRoom = false;
                }
                if (args[0] == "MasterOpeRoom")
                {
                    GlobalValues.MasterMode = true;
                    GlobalValues.MasterOpeRoom = true;
                }
            }

            //
            // Nlog用
            // logs フォルダの存在確認と作成
            //
            string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            if (!Directory.Exists(logDir))
            {
                try
                {
                    Directory.CreateDirectory(logDir);
                    logger.Info($"ログフォルダを作成しました: {logDir}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ログフォルダの作成に失敗しました: {ex.Message}");
                    return;
                }
            }




            var xlsxFiles = PickListHelper.GetPlFilePathsFromXml();

            var movedFiles = new List<string>();

            if (PickListHelper.AllFilesMatchPattern(xlsxFiles))
            {
                if (PickListHelper.PrepareOriginalXlsxFolder(out string folderPath))
                {
                    logger.Info($"OriginalXLSX folder created at: {folderPath}");
                }
                else
                {
                    // エラーはすでに表示されているので、終了
                    return;
                }

                //
                // ファイルをOriginalXLSXへ（PDF, XLSX)移動してファイルリストを得る
                // XLSXファイルのみのリストになる
                //
                movedFiles = PickListHelper.MoveMatchingFilesToOriginalXlsxFolder();
                logger.Info("移動されたファイルの絶対パス一覧:");
                foreach (var path in movedFiles)
                {
                    logger.Info(path);
                }


            }
            else
            {
                logger.Error("ファイル名が条件に一致しないものがあります。処理を中止します。");
                return;
            }

            string templatePath = TemplateHelper.GetTemplateFilePath();
            logger.Info("Template Path: " + templatePath);

            //
            // Original ファイルに 申送りシートと、残存防止シートを出力 ExcelCreator12使用
            // 入れたxlsxは workの下におかれる
            // Insertedfilelistには、申残のxlsxリストが変える。
            //
            //
            List<string> Insertedfilelist = PickListHelper.InsertTemplateSheetsIntoFiles(movedFiles, templatePath);

            //
            // 申し送りシートを挿入作成
            // ExcelCreatorを使用 ざっくり5%
            //
            MoushiokuriHelper.InsertMoushiokuriItems(Insertedfilelist);

            //
            // 残存防止シートを挿入作成
            // ExcelCreatorを使用 ざっくり25%
            //
            ZanzonboushiHelper.InsertZanzonboushiItems(Insertedfilelist);

            //
            // 両面対応
            //
            DoubleSidedFormatHelper.InsertBlankPages(Insertedfilelist);

            //
            // PDFを作成する
            //
            PDFoutputHelper.PDFfilesOutput(Insertedfilelist);

        }
    }
}
