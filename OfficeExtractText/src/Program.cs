using Microsoft.Extensions.CommandLineUtils;
using OfficeExtractText.Properties;
using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace OfficeExtractText
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en");
            if (CultureInfo.InstalledUICulture.TwoLetterISOLanguageName == "ja")
            {
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("ja");
            }

            //Definition of application command arguments
            var app = new CommandLineApplication(throwOnUnexpectedArg: false);
            app.Name = "OfficeExtractText.exe";
            app.Description = Resource.APP_DESCRIPTION;
            var targetPathArgument = app.Argument("<target file or directory>", Resource.TARGET_PATH_ARGUMENT);
            var outputOption = app.Option("-o|--output", Resource.OUTPUT_OPTION, CommandOptionType.SingleValue);
            var subDirOption = app.Option("-s|--subdir", Resource.SUBDIR_OPTION, CommandOptionType.NoValue);
            var excelOption = app.Option("-e|--excel", Resource.EXCEL_OPTION, CommandOptionType.NoValue);
            var wordOption = app.Option("-w|--word", Resource.WORD_OPTION, CommandOptionType.NoValue);
            var powerPointOption = app.Option("-p|--powerpoint", Resource.POWERPOINT_OPTION, CommandOptionType.NoValue);
            var nologOption = app.Option("--no-log", Resource.NOLOG_OPTION, CommandOptionType.NoValue);

            app.OnExecute(() =>
            {
                //Options setting.
                if (targetPathArgument.Value == null || !outputOption.HasValue())
                {
                    app.ShowHelp();
                    return 3;
                }
                //Init logger.
                ConsoleLogger.InitLogger(!nologOption.HasValue(), !nologOption.HasValue(), true);

                //Init options.
                var extractExcel = excelOption.HasValue();
                var extractWord = wordOption.HasValue();
                var extractPowerPoint = powerPointOption.HasValue();
                if (!extractExcel && !extractWord && !extractPowerPoint)
                {
                    extractExcel = true;
                    extractWord = true;
                    extractPowerPoint = true;
                }
                var options = new OfficeTextExporterOption();
                options.ExtractExcel = extractExcel;
                options.ExtractWord = extractWord;
                options.ExtractPowerPoint = extractPowerPoint;
                options.ExtractSubDir = subDirOption.HasValue();
                options.OutputDir = Path.GetFullPath(outputOption.Value());
                try
                {
                    //Execute exporter.
                    var targetPath = Path.GetFullPath(targetPathArgument.Value);
                    ConsoleLogger.WriteLog(String.Format(Resource.EXTRACTING_STARTED_PATH_ARG0, targetPath));
                    var exporter = new OfficeTextExporter(targetPath, options);
                    exporter.Execute();
                    if (exporter.HasWarning())
                    {
                        ConsoleLogger.WriteLog(Resource.EXTRACTING_FINISHED_WITH_WARNING);
                        return 2;
                    }
                    else
                    {
                        ConsoleLogger.WriteLog(Resource.EXTRACTING_FINISHED_SUCCESSFULLY);
                        return 0;
                    }
                }
                catch (ApplicationException e)
                {
                    ConsoleLogger.WriteError(String.Format(Resource.EXTRACTING_FINISHED_WITH_ERROR_DETAILS_ARG0, e.Message));
                    return 1;
                }
                catch (Exception e)
                {
                    ConsoleLogger.WriteError(String.Format(Resource.EXTRACTING_FINISHED_WITH_ERROR_DETAILS_ARG0, e.ToString()));
                    return 1;
                }
            });

            try
            {
                app.Execute(args);
            }
            catch (CommandParsingException)
            {
                app.ShowHelp();
                Environment.Exit(2);
            }
        }
    }
}
