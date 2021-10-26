using CommandLine;
using Multi_Language_Xaml.Generator.Models;
using Npoi.Mapper;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Multi_Language_Xaml.Generator
{
    public class Program
    {
        public static async Task Main(string[] args)
            => await Parser.Default.ParseArguments<Options>(args).WithParsedAsync(RunAsync);

        public static async Task RunAsync(Options opts)
        {
            bool isSuccess = CheckExcel(opts.ExcelFile, opts.Key, opts.Language);
            if (!isSuccess)
                return;
            foreach(string language in opts.Language)
            {
                LanguageDto[] languages = GetLanguage(opts.ExcelFile, opts.Key, language);
                await SaveXaml(languages, language, opts.ExportPath);
                Console.WriteLine($"{language}.xaml done.");
            }

            Console.Write("\r\nPress any key to continue...");
            Console.ReadKey();
        }

        private static (bool, IWorkbook) LoadExcel(string fileName)
        {
            IWorkbook workbook;
            try
            {
                workbook = new XSSFWorkbook(fileName);
                return (true, workbook);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return (false, null);
            }
        }

        private static bool CheckExcel(string excelFile, string key, IEnumerable<string> languages)
        {
            (bool isOpen, IWorkbook workbook) = LoadExcel(excelFile);
            if (!isOpen)
                return false;
            ISheet sheet = workbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);
            if (!headerRow.Cells.Any(item => item.ToString() == key))
            {
                Console.WriteLine("Key is not found.");
                return false;
            }
            foreach (string language in languages)
            {
                if (!headerRow.Cells.Any(item => item.ToString() == language))
                {
                    Console.WriteLine($"{language} is not found.");
                    return false;
                }
            }
            workbook.Close();
            Mapper mapper = new(excelFile);
            mapper.Map<KeyDto>(key, s => s.Key);
            IEnumerable<string> keys = mapper.Take<KeyDto>().Select(item => item.Value.Key);
            if (keys.GroupBy(item => item).Where(item => item.Count() > 1).Any())
            {
                Console.WriteLine("There are duplicate keys ");
                return false;
            }
            return true;
        }

        private static LanguageDto[] GetLanguage(string excelFile, string key, string language)
        {
            Mapper mapper = new(excelFile);
            mapper.Map<LanguageDto>(key, s => s.Key);
            mapper.Map<LanguageDto>(language, s => s.Language);
            return mapper.Take<LanguageDto>().Select(item => item.Value).ToArray();
        }

        private static async Task SaveXaml(LanguageDto[] languages, string language, string savePath)
        {
            List<string> xamlContent = new();
            xamlContent.Add("<ResourceDictionary xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"");
            xamlContent.Add("                    xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"");
            xamlContent.Add("                    xmlns:system=\"clr-namespace:System;assembly=mscorlib\">");
            foreach(LanguageDto languageDto in languages)
                xamlContent.Add($"    <system:String x:Key=\"{languageDto.Key}\" xml:space=\"preserve\">{languageDto.Language}</system:String>");
            xamlContent.Add("</ResourceDictionary>");

            if (savePath is not null)
            {
                if (!Directory.Exists(savePath))
                    Directory.CreateDirectory(savePath);
                string parh = Path.Combine(savePath, $"{language}.xaml");
                await File.WriteAllLinesAsync(parh, xamlContent);
            }
            else
                await File.WriteAllLinesAsync($"{language}.xaml", xamlContent);
        }
    }
}
