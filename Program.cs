using Newtonsoft.Json;
using OfficeOpenXml;

public class Config
{
    public double AnnualInterestRate { get; set; }
    public string? inputFile { get; set; }
    public string? outputFile { get; set; }
    public bool OverwriteExistingFile { get; set; }

    public void Print()
    {
        Console.WriteLine("CONFIG:");
        Console.WriteLine("- inputFile: " + inputFile);
        Console.WriteLine("- outputFile: " + outputFile);
        Console.WriteLine("- OverwriteExistingFile: " + OverwriteExistingFile);
        Console.WriteLine();
    }
}

public class InputModel
{
    public string Description { get; set; }
    public DateTime Date { get; set; }
    public double Amount { get; set; }
}

public class OutputModel
{
    public string Description { get; set; }
    public string Date { get; set; }
    public double Amount { get; set; }
}

class Program
{
    public static Config Config;
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Config = ReadConfig("config.json");

        if (string.IsNullOrEmpty(Config.inputFile))
        {
            Console.Write("Nazwa pliku wejsciowego: ");
            Config.inputFile = Console.ReadLine();
        }
        if (string.IsNullOrEmpty(Config.outputFile))
        {
            Console.Write("Nazwa pliku wyjściowego: ");
            Config.outputFile = Console.ReadLine();
        }

        Config.Print();

        List<InputModel> data = ReadData(Config.inputFile);
        List<OutputModel> results = CalculateInterest(data, Config.AnnualInterestRate / 100);
        SaveData(Config.outputFile, results);
    }

    static Config ReadConfig(string configPath)
    {
        using (StreamReader reader = new StreamReader(configPath))
        {
            string json = reader.ReadToEnd();
            return JsonConvert.DeserializeObject<Config>(json);
        }
    }

    static List<InputModel> ReadData(string filePath)
    {
        var data = new List<InputModel>();
        string fileExtension = Path.GetExtension(filePath).ToLower();

        if (fileExtension == ".csv")
        {
            // Wczytaj dane z pliku CSV
            using (var reader = new StreamReader(filePath))
            {
                var headerLine = reader.ReadLine(); // Skip header line
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');

                    var inputData = new InputModel
                    {
                        Description = values[0],
                        Date = DateTime.Parse(values[1]),
                        Amount = double.Parse(values[2])
                    };
                    data.Add(inputData);
                }
            }
        }
        else if (fileExtension == ".xlsx" || fileExtension == ".xls")
        {
            // Wczytaj dane z pliku Excel
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int row = 2; // Start from the second row (assuming the first row is the header)

                while (worksheet.Cells[row, 1].Value != null)
                {
                    var inputData = new InputModel
                    {
                        Description = worksheet.Cells[row, 1].Text,
                        Date = DateTime.Parse(worksheet.Cells[row, 2].Text),
                        Amount = double.Parse(worksheet.Cells[row, 3].Text)
                    };
                    data.Add(inputData);
                    row++;
                }
            }
        }
        else
        {
            throw new InvalidOperationException("Unsupported file type. Please use a .csv or .xlsx file.");
        }

        return data;
    }

    static List<OutputModel> CalculateInterest(List<InputModel> data, double annualInterestRate)
    {
        var results = new List<OutputModel>();
        double dailyInterestRate = annualInterestRate / 365;

        foreach (var entry in data)
        {
            double amount = entry.Amount;
            DateTime currentDate = entry.Date;
            DateTime presentDate = DateTime.Now;

            while (currentDate.AddDays(30) < presentDate)
            {
                DateTime nextDate = currentDate.AddDays(30);
                double numberOfDays = (nextDate - currentDate).TotalDays;
                double accruedInterest = amount * dailyInterestRate * numberOfDays;

                amount += accruedInterest;

                results.Add(new OutputModel
                {
                    Description = entry.Description,
                    Date = nextDate.ToString("yyyy-MM-dd"),
                    Amount = Math.Round(amount, 2)
                });

                currentDate = nextDate;
            }
        }

        return results;
    }


    static void SaveData(string filePath, List<OutputModel> wyniki)
    {
        if (File.Exists(filePath) && !Config.OverwriteExistingFile)
            filePath = filePath.Insert(filePath.LastIndexOf('.'), DateTime.Now.ToString("yyyy-MM-dd HH-MM-ss"));


        string fileExtension = Path.GetExtension(filePath).ToLower();

        if (fileExtension == ".csv")
        {
            // Zapisz wyniki do pliku CSV
            using (var writer = new StreamWriter(filePath))
            {
                writer.WriteLine("Opis,Data,Kwota");
                foreach (var wynik in wyniki)
                {
                    writer.WriteLine($"{wynik.Description},{wynik.Date},{wynik.Amount}");
                }
            }
        }
        else if (fileExtension == ".xlsx" || fileExtension == ".xls")
        {
            // Zapisz wyniki do pliku Excel
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Wyniki");
                worksheet.Cells[1, 1].Value = "Opis";
                worksheet.Cells[1, 2].Value = "Data";
                worksheet.Cells[1, 3].Value = "Kwota";

                int row = 2;
                foreach (var wynik in wyniki)
                {
                    worksheet.Cells[row, 1].Value = wynik.Description;
                    worksheet.Cells[row, 2].Value = wynik.Date;
                    worksheet.Cells[row, 3].Value = wynik.Amount;
                    row++;
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
        else
        {
            throw new InvalidOperationException("Unsupported file type. Please use a .csv or .xlsx file.");
        }

        Console.WriteLine("Wyniki zapisano w pliku: " + filePath);
    }
}
