using CommandLine;
using System.Collections.Generic;

namespace Multi_Language_Xaml.Generator.Models
{
    public class Options
    {
        [Option('i', "input", Required = true, HelpText = "紀錄多國語言的 Excel")]
        public string ExcelFile { get; set; }
        [Option('k', "Key", Required = true, HelpText = "在 Excel 中多國語言的主鍵欄位")]
        public string Key { get; set; }
        [Option('l', "Language ", Required = true, HelpText = "在 Excel 中多國語言的語言資料欄位")]
        public IEnumerable<string> Language { get; set; }
        [Option('o', "output", Required = false, HelpText = "Xaml 輸出路徑")]
        public string ExportPath { get; set; }
    }
}
