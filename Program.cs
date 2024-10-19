using Newtonsoft.Json;
using OfficeOpenXml;

namespace Calculator;

public class Config(double annualInterestRate, bool overwriteExistingFile)
{
    public double AnnualInterestRate { get; } = annualInterestRate;
    public string? InputFile { get; set; }
    public string? OutputFile { get; set; }
    public bool OverwriteExistingFile { get; } = overwriteExistingFile;

    public void Print()
    {
        Console.WriteLine("CONFIG:");
        Console.WriteLine("- inputFile: " + InputFile);
        Console.WriteLine("- outputFile: " + OutputFile);
        Console.WriteLine("- OverwriteExistingFile: " + OverwriteExistingFile);
        Console.WriteLine();
    }
}

/// <summary>
/// Expects input file in Excel format containing a header describing columns:
/// description, date, amount
/// And a list of calculations below the header whereas each
/// date indicates a start date for calculating monthly interests
/// </summary>
internal class InputModel
{
    public string? Description { get; init; }
    public DateTime Date { get; init; }
    public double Amount { get; init; }
}

internal class OutputModel(string date)
{
    public string? Description { get; init; }
    public string Date { get; init; } = date;
    public double Amount { get; init; }
}

abstract class Program
{
    // TODO: must extend the system by reading interest rates from another file containing columns:
    // Date, Interest - so that interest rates are effective from the specified date (inclusive)
    // and the calculation must take into account every change of interest for each calculation
    private static Config? _config;
    
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        _config = ReadConfig("config.json");

        if (string.IsNullOrEmpty(_config?.InputFile))
        {
            Console.Write("Nazwa pliku wejsciowego: ");
            if (_config != null) _config.InputFile = Console.ReadLine();
        }
        if (string.IsNullOrEmpty(_config?.OutputFile))
        {
            Console.Write("Nazwa pliku wyjściowego: ");
            if (_config != null) _config.OutputFile = Console.ReadLine();
        }

        _config?.Print();

        if (_config?.InputFile == null) return;
        List<InputModel> data = ReadData(_config.InputFile);
        List<OutputModel> results = CalculateInterest(data, _config.AnnualInterestRate / 100);
        if (_config.OutputFile != null) SaveData(_config.OutputFile, results);
    }

    private static Config? ReadConfig(string configPath)
    {
        using StreamReader reader = new(configPath);
        string json = reader.ReadToEnd();
        return JsonConvert.DeserializeObject<Config>(json);
    }

    private static List<InputModel> ReadData(string filePath)
    {
        List<InputModel> data = [];
        string fileExtension = Path.GetExtension(filePath).ToLower();

        switch (fileExtension)
        {
            case ".csv":
            {
                using StreamReader reader = new(filePath);
                string? headerLine = reader.ReadLine(); // Skip header line
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    string?[]? values = line?.Split(';');

                    InputModel inputData = new()
                    {
                        Description = values?[0],
                        Date = DateTime.Parse(values?[1]),
                        Amount = double.Parse(values?[2])
                    };
                    data.Add(inputData);
                }

                break;
            }
            case ".xlsx":
            case ".xls":
            {
                // Wczytaj dane z pliku Excel
                using ExcelPackage package = new(new FileInfo(filePath));
                ExcelWorksheet? worksheet = package.Workbook.Worksheets[0];
                int row = 2; // Start from the second row (assuming the first row is the header)

                while (worksheet.Cells[row, 1].Value != null)
                {
                    InputModel inputData = new()
                    {
                        Description = worksheet.Cells[row, 1].Text,
                        Date = DateTime.Parse(worksheet.Cells[row, 2].Text),
                        Amount = double.Parse(worksheet.Cells[row, 3].Text)
                    };
                    data.Add(inputData);
                    row++;
                }

                break;
            }
            default:
                throw new InvalidOperationException("Unsupported file type. Please use a .csv or .xlsx file.");
        }

        return data;
    }

    static List<OutputModel> CalculateInterest(List<InputModel> data, double annualInterestRate)
    {
        List<OutputModel> results = [];
        double dailyInterestRate = annualInterestRate / 365;

        foreach (InputModel entry in data)
        {
            double amount = entry.Amount;
            DateTime currentDate = entry.Date;
            DateTime presentDate = DateTime.Now;

            while (currentDate.AddMonths(1) < presentDate)
            {
                DateTime nextDate = currentDate.AddMonths(1);
                double numberOfDays = (nextDate - currentDate).TotalDays;
                double accruedInterest = amount * dailyInterestRate * numberOfDays;

                amount += accruedInterest;

                results.Add(new OutputModel(nextDate.ToString("yyyy-MM-dd"))
                {
                    Description = entry.Description,
                    Amount = Math.Round(amount, 2)
                });

                currentDate = nextDate;
            }
        }

        return results;
    }


    static void SaveData(string filePath, List<OutputModel> wyniki)
    {
        if (File.Exists(filePath) && _config is { OverwriteExistingFile: false })
            filePath = filePath.Insert(filePath.LastIndexOf('.'), DateTime.Now.ToString("yyyy-MM-dd HH-MM-ss"));
        
        string fileExtension = Path.GetExtension(filePath).ToLower();

        switch (fileExtension)
        {
            case ".csv":
            {
                // Zapisz wyniki do pliku CSV
                using StreamWriter writer = new(filePath);
                writer.WriteLine("Opis,Data,Kwota");
                foreach (OutputModel wynik in wyniki)
                {
                    writer.WriteLine($"{wynik.Description},{wynik.Date},{wynik.Amount}");
                }

                break;
            }
            case ".xlsx":
            case ".xls":
            {
                // Zapisz wyniki do pliku Excel
                using ExcelPackage package = new();
                ExcelWorksheet? worksheet = package.Workbook.Worksheets.Add("Wyniki");
                worksheet.Cells[1, 1].Value = "Opis";
                worksheet.Cells[1, 2].Value = "Data";
                worksheet.Cells[1, 3].Value = "Kwota";

                int row = 2;
                foreach (OutputModel wynik in wyniki)
                {
                    worksheet.Cells[row, 1].Value = wynik.Description;
                    worksheet.Cells[row, 2].Value = wynik.Date;
                    worksheet.Cells[row, 3].Value = wynik.Amount;
                    row++;
                }

                package.SaveAs(new FileInfo(filePath));

                break;
            }
            default:
                throw new InvalidOperationException("Unsupported file type. Please use a .csv or .xlsx file.");
        }

        Console.WriteLine("Wyniki zapisano w pliku: " + filePath);
    }
}