using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Resources;
using System.Resources.NetStandard;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;

internal class Program
{
    private static void Main(string[] args)
    {
        string baseDirectory = @"C:\Users\joao.oliveira\Downloads\BusinessOne\AgriBusinessAlgodoeira\Apontamentos\BaixaFardinhos";

        // Busca todos os arquivos .resx no diretório base e suas subpastas
        List<string> resxFiles = GetResxFiles(baseDirectory);

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        foreach (string resxFile in resxFiles)
        {
            // Cria o arquivo .xlsx com o nome da pasta do arquivo .resx
            string xlsxFile = Path.Combine(
                Path.GetDirectoryName(resxFile),
                $"{new DirectoryInfo(Path.GetDirectoryName(resxFile)).Name}.xlsx"
            );

            using (ResXResourceReader reader = new ResXResourceReader(resxFile))
            {
                Dictionary<string, string> values = new Dictionary<string, string>();
                foreach (DictionaryEntry entry in reader)
                {
                    values.Add(entry.Key.ToString(), entry.Value.ToString());
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(xlsxFile)))
                {
                    if (File.Exists(xlsxFile))
                    {
                        File.Delete(xlsxFile);
                    }

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Aba1");

                    int row = 1;
                    foreach (var item in values)
                    {
                        worksheet.Cells[row, 1].Value = item.Key.ToString();
                        worksheet.Cells[row, 2].Value = item.Value.ToString();
                        row++;
                    }
                    package.Save();
                }
            }
        }
    }

    private static List<string> GetResxFiles(string directory)
    {
        List<string> resxFiles = new List<string>();

        foreach (string file in Directory.GetFiles(directory, "lang.en-US.resx"))
        {
            resxFiles.Add(file);
        }

        foreach (string subDirectory in Directory.GetDirectories(directory))
        {
            resxFiles.AddRange(GetResxFiles(subDirectory));
        }

        return resxFiles;
    }
}
