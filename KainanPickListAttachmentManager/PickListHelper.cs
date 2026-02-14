using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using AdvanceSoftware.ExcelCreator;
using NLog;





namespace KainanPickListAttachmentManager.Helpers
{

    public class PickListHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static AdvanceSoftware.ExcelCreator.Creator creator10 = new AdvanceSoftware.ExcelCreator.Creator();
        private static AdvanceSoftware.ExcelCreator.Creator creator11 = new AdvanceSoftware.ExcelCreator.Creator();

        private static void creator10_Error(object sender,
            AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator10: Error processing file");
            logger.Error("creator2: " + creator11.ErrorMessage);
        }

        private static void creator11_Error(object sender,
            AdvanceSoftware.ExcelCreator.CreatorErrorEventArgs e)
        {
            // ここにエラー発生時の処理を追加します。
            logger.Error($"ExcelCreator12 object creator11: Error processing file");
            logger.Error("creator2: " + creator11.ErrorMessage);
        }


        public static List<string> GetXlsxFilePaths()
        {
            var filePaths = new List<string>();
            string folderPath = System.Configuration.ConfigurationManager.AppSettings["PickListFolderPath"];

            if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
            {
                string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                filePaths.AddRange(files);
            }

            return filePaths;
        }

        public static List<string> GetPlFilePathsFromXml()
        {
            var filePaths = new List<string>();
            string folderPath;
            if (GlobalValues.MasterMode && (GlobalValues.MasterOpeRoom ) == false)
            {
                folderPath = ConfigurationManager.AppSettings["MasterPLFolderPath"];
            }
            else
            {
                folderPath = ConfigurationManager.AppSettings["PickListFolderPath"];
            }

            string xmlPath = Path.Combine(folderPath, "EOR_PLFiles.xml");

            if (!File.Exists(xmlPath))
            {
                logger.Error($"XMLファイルが見つかりません: {xmlPath}");
                return filePaths;
            }




            if (GlobalValues.MasterMode)
            {
                try
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(xmlPath);

                    XmlNodeList plfileNodes = doc.SelectNodes("//opemethodfile");
                    foreach (XmlNode node in plfileNodes)
                    {
                        string fname = node.Attributes["filename"]?.Value;
                        if (!string.IsNullOrEmpty(fname))
                        {
                            string fileName = fname + ".xlsx";
                            string fullPath = Path.Combine(folderPath, fileName);
                            filePaths.Add(fullPath);
                            logger.Info(fullPath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error($"XMLの読み込み中にエラーが発生しました: {ex.Message}");
                }
            }
            else
            { 
                try
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(xmlPath);

                    XmlNodeList plfileNodes = doc.SelectNodes("//plfile");
                    foreach (XmlNode node in plfileNodes)
                    {
                        string id = node.Attributes["id"]?.Value;
                        if (!string.IsNullOrEmpty(id))
                        {
                            string fileName = id + "手術準備表.xlsx";
                            string fullPath = Path.Combine(folderPath, fileName);
                            filePaths.Add(fullPath);
                            logger.Info(fullPath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error($"XMLの読み込み中にエラーが発生しました: {ex.Message}");
                }

            }
            return filePaths;
        }




        public static bool AllFilesMatchPattern(List<string> filePaths)
        {
            var pattern = new Regex(@".*手術準備表\.xlsx$");
            foreach (var path in filePaths)
            {
                string fileName = Path.GetFileName(path);
                if (!pattern.IsMatch(fileName))
                {
                    return false;
                }
            }
            return true;
        }



        public static bool PrepareOriginalXlsxFolder(out string folderPath)
        {

            string xlsxPath;
            if (GlobalValues.MasterMode && (GlobalValues.MasterOpeRoom) == false)
            {
                xlsxPath = System.Configuration.ConfigurationManager.AppSettings["MasterPLFolderPath"];
            }
            else
            {
                xlsxPath = System.Configuration.ConfigurationManager.AppSettings["PickListFolderPath"];

            }
            
            folderPath = Path.Combine(xlsxPath, "OriginalXLSX");

            try
            {
                if (Directory.Exists(folderPath))
                {
                    Directory.Delete(folderPath, true);
                }

                Directory.CreateDirectory(folderPath);
                return true;
            }
            catch (IOException)
            {
                logger.Error($"{folderPath} フォルダは編集中です。");
                return false;
            }
            catch (Exception)
            {
                logger.Error($"{folderPath} フォルダが作成できません。");
                return false;
            }
        }

        public static List<string> MoveMatchingFilesToOriginalXlsxFolder()
        {
            var movedFilePaths = new List<string>();

            string sourceFolder;
            if (GlobalValues.MasterMode && (GlobalValues.MasterOpeRoom) == false)
            {
                sourceFolder = ConfigurationManager.AppSettings["MasterPLFolderPath"];
            }
            else
            {
                sourceFolder = ConfigurationManager.AppSettings["PickListFolderPath"];
            }

            string originalFolder = Path.Combine(sourceFolder, "OriginalXLSX");

            var logger = LogManager.GetCurrentClassLogger();

            // ファイル名パターン
            var pattern = new Regex(@".*手術準備表\.xlsx$");

            foreach (var file in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                string fileName = Path.GetFileName(file);
                if (pattern.IsMatch(fileName))
                {
                    string destPath = Path.Combine(originalFolder, fileName);
                    try
                    {
                        File.Move(file, destPath);
                        movedFilePaths.Add(Path.GetFullPath(destPath));
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"XLSXファイルの移動に失敗しました: {file}");
                        // 移動できなかったファイルはリストに追加しない
                    }
                }
            }

            // PDF移動
            // ファイル名パターン
            var movedFilePaths2 = new List<string>();
            var pattern2 = new Regex(@".*手術準備表\.pdf$");

            foreach (var file in Directory.GetFiles(sourceFolder, "*.pdf"))
            {
                string fileName = Path.GetFileName(file);
                if (pattern2.IsMatch(fileName))
                {
                    string destPath = Path.Combine(originalFolder, fileName);
                    try
                    {
                        File.Move(file, destPath);
                        movedFilePaths2.Add(Path.GetFullPath(destPath));
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"PDFファイルの移動に失敗しました: {file}");
                        // 移動できなかったファイルはリストに追加しない
                    }
                }
            }

            return movedFilePaths;
        }


        //
        //
        // movedFiles（OriginalXLSXフォルダの下のXLSXファイルリスト）と
        // 申し送り残存防止のテンプレートファイルを得る
        //
        //
        public static List<string> InsertTemplateSheetsIntoFiles(List<string> movedFiles, string templatePath)
        {
            var logger = LogManager.GetCurrentClassLogger();

            string folderPath;
            if (GlobalValues.MasterMode && (GlobalValues.MasterOpeRoom) == false)
            {
                folderPath = ConfigurationManager.AppSettings["MasterPLFolderPath"];
            }
            else
            {
                folderPath = ConfigurationManager.AppSettings["PickListFolderPath"];
            }


            var progressForm = new ProgressForm();
            progressForm.Show();

            System.Reflection.Assembly myAssembly10 = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense10 = myAssembly10.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator10.SetLicense(tempStreamLicense10);

            System.Reflection.Assembly myAssembly11 = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream tempStreamLicense11 = myAssembly11.GetManifestResourceStream("KainanPickListAttachmentManager.AdvanceSoftware.ExcelCreator12.License.xml");
            creator11.SetLicense(tempStreamLicense11);

            List<string> return_filelist = new List<string>();



            try
            {


                int i = 0;
                foreach (var filePath in movedFiles)
                {


                    try
                    {
                        //
                        // プログレスバー処理
                        //
                        progressForm.UpdateProgress(i + 1, movedFiles.Count, $"処理中: {Path.GetFileName(filePath)}");
                        i++;


                        //
                        // XLSXをPLを付けて保存
                        //
                        string PLfileName = Path.GetFileNameWithoutExtension(filePath) + "ＰＬ" +Path.GetExtension(filePath);
                        string savePath = Path.Combine(folderPath, PLfileName);


                        creator10.OpenBook(savePath, filePath);
                        creator10.CloseBook(true);

                        string TEMPLfileName = Path.GetFileNameWithoutExtension(filePath) + "申残" + Path.GetExtension(templatePath);
                        string save11Path = Path.Combine(folderPath, TEMPLfileName);


                        //
                        // 申送り、残存防止テンプレートを申残をつけて保存
                        //
                        creator11.OpenBook(save11Path, templatePath);
                        creator11.CloseBook(true);


                        return_filelist.Add(save11Path);

                        logger.Info($"XLSXをPLを付けて保存、テンプレートを申残をつけて保存しました: {savePath}");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"ファイル処理中にエラーが発生しました: {filePath}");

                    }
                }


            }
            catch (Exception ex)
            {
                logger.Error(ex, "Excel の起動に失敗しました。");
            }
            finally
            {
                progressForm.Close();

            }

            return return_filelist;
        }


    }
}