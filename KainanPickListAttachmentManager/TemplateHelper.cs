using System;
using System.IO;
using NLog;

namespace KainanPickListAttachmentManager.Helpers
{
    public class TemplateHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public static string GetTemplateFilePath()
        {
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;

            string templateFileName = System.Configuration.ConfigurationManager.AppSettings["PickListAttachmentTemplate"];
            string templatePath = Path.Combine(exeDirectory, templateFileName);


            if (File.Exists(templatePath))
            {
                logger.Info("テンプレート読み込み:" + templatePath);
                return templatePath;
            }
            else
            {
                logger.Error("PickListAttachmentTemplateファイルがありません:" + templatePath);
                return null;
            }

        }
    }
}
