namespace TaxesHelper
{
    using System;
    using System.Collections.Generic;
    using System.CommandLine;
    using System.CommandLine.Invocation;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using System.Web;

    using OfficeOpenXml;

    public static class Program
    {
        private static readonly HttpClient HttpClient = new();

        public static async Task Main(string[] args)
        {
            var rootCommand = new RootCommand();
            rootCommand.AddCommand(CreateSellsCommand());
            rootCommand.AddCommand(CreateHoldingsCommand());
            rootCommand.AddCommand(CreateDividendCommand());

            await rootCommand.InvokeAsync(args);
        }

        private static Command CreateSellsCommand()
        {
            var command = new Command(
            "sells", "This command generates the output for Application №5.");

            var fileOption = new Option<string>("--file")
            {
                IsRequired = true,
                Description = @"Path to the Etrade's excel file for sellable stocks.
Go to Etrade > My Account > Gains & Losses > Download Expanded and save the file locally.
Note: you must filter by year!"
            };

            var pricesOption = new Option<decimal[]>("--sell-prices")
            {
                IsRequired = true,
                Description = @"List of the sell order prices that are in the provided file.
Etrade provides an 'Adjusted price' which includes their commission.
To find out your sell prices, go to Etrade > My Account > Orders > Execution price.
Note: each price must be provided as a separate argument e.g. --sell-prices 100 120"
            };

            command.AddOption(fileOption);
            command.AddOption(pricesOption);
            command.Handler = CommandHandler.Create(
                async (string file, decimal[] sellPrices) => await AnalyzeSells(file, sellPrices));

            return command;
        }

        private static Command CreateHoldingsCommand()
        {
            var fileOption = new Option<string>("--file")
            {
                IsRequired = true,
                Description = @"Path to the Etrade's excel file for sellable stocks.
Go to Etrade > Stock Plan > View By Status > Download Expanded and save the file locally."
            };

            var interactiveOption = new Option<bool>("--interactive")
            {
                IsRequired = false,
                Description = @"If true the program prints 1 line and waits for user input before printing the next.
Useful if you're filling the application. If false (default), dumps the whole output in one go."
            };

            var command = new Command(
            "holdings",
            "This command generates the output for Application №8.");

            command.AddOption(fileOption);
            command.AddOption(interactiveOption);

            command.Handler = CommandHandler.Create(
                async (string file, bool? interactive) =>
                    await AnalyzeHoldings(file, interactive.HasValue && interactive.Value));

            return command;
        }

        private static Command CreateDividendCommand()
        {
            var dividendArgument = new Argument("dividend")
            {
                Arity = ArgumentArity.ExactlyOne,
                Description = @"The un-taxed dividend that you received.
Go to Etrade > Holdings > Other Holdings & expand the cash section and find the oldest entry."
            };

            var command = new Command("dividend", "This command generates the output for Application №8, Part IV.");
            command.AddArgument(dividendArgument);

            command.Handler = CommandHandler.Create(
                async (decimal dividend) => await AnalyzeDividend(dividend));

            return command;
        }

        private static async Task AnalyzeDividend(decimal dividendUsd)
        {
            var rate = await GetBnbRate(new DateTime(2021, 04, 11));

            Console.WriteLine($"1 USD = {rate} лв.");
            Console.WriteLine();

            var dividendBgn = rate * dividendUsd;

            var usdTaxesPayedInUs = dividendUsd * 0.1M * 0.3949M;
            var bgnTaxesPayedInUs = rate * usdTaxesPayedInUs;

            var taxCredit = dividendBgn * 0.05M;

            var taxToPay = taxCredit - bgnTaxesPayedInUs;

            Console.WriteLine($"Наименование на лицето, изплатило дохода: VMware, Inc.");
            Console.WriteLine($"Държава: САЩ");
            Console.WriteLine($"Код вид доход: 8141");
            Console.WriteLine($"Код за прилагане на метод за избягване на двойното данъчно облагане: 1");
            Console.WriteLine($"Брутен размер на дохода: {dividendBgn:.00}");
            Console.WriteLine($"Документално доказана цена на придобиване: 0.00");
            Console.WriteLine($"Положителна разлика между колона 6 и колона 7: 0.00");
            Console.WriteLine($"Платен данък в чужбина: {bgnTaxesPayedInUs:.00}");
            Console.WriteLine($"Допустим размер на данъчния кредит: {taxCredit:.00}");
            Console.WriteLine($"Размер на признатия данъчен кредит: {taxCredit:.00}");
            Console.WriteLine($"Дължим данък, подлежащ на внасяне: {taxToPay:.00}");
        }

        private static async Task AnalyzeHoldings(string file, bool interactive)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var culture = CultureInfo.CreateSpecificCulture("en-us");

            using var package = new ExcelPackage(new FileInfo(file));
            var worksheet = package.Workbook.Worksheets[0];

            var totalRows = worksheet.Dimension.Rows - 2;
            Console.WriteLine($"Total rows: {totalRows}");
            var delimiter = "    ";
            Console.WriteLine($"№{delimiter} Брой {delimiter} Дата на придобиване {delimiter} Стойност: {delimiter} $ Стойност: лв.");
            for (int row = 2; row <= worksheet.Dimension.Rows - 1; row++)
            {
                var acquireDate = DateTime.Parse(worksheet.Cells[row, 4].GetValue<string>(), culture);
                var quantity = worksheet.Cells[row, 5].GetValue<int>();
                var usdStockPrice = decimal.Parse(worksheet.Cells[row, 28].GetValue<string>().Substring(1));
                var rate = await GetBnbRate(acquireDate);
                var totalUsd = quantity * usdStockPrice;
                var totalBgn = totalUsd * rate;

                var sb = new StringBuilder();
                var printValues = new List<string>
                {
                    (row - 1).ToString(),
                    quantity.ToString(),
                    acquireDate.ToString("dd.MM.yyyy"),
                    $"{totalUsd:.00} $",
                    $"{totalBgn:.00} лв.",
                };

                printValues.ForEach(x => sb.AppendFormat("{0}{1}", x, delimiter));
                if (interactive)
                {
                    Console.ReadLine();
                }
                Console.WriteLine(sb);
            }
        }

        private static async Task AnalyzeSells(string file, IReadOnlyList<decimal> sellPrices)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var culture = CultureInfo.CreateSpecificCulture("en-us");

            using var package = new ExcelPackage(new FileInfo(file));
            var worksheet = package.Workbook.Worksheets[0];

            if (sellPrices.Count != worksheet.Dimension.End.Row - 2)
            {
                throw new ArgumentException($"Expected to have {worksheet.Dimension.End.Row - 2} sell prices, not {sellPrices.Count}");
            }
            var bgnBenefit = 0M;
            for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
            {
                var quantity = worksheet.Cells[row, 4].GetValue<int>();

                var buyDate = DateTime.Parse(worksheet.Cells[row, 5].GetValue<string>(), culture);
                var buyUsdPricePerStock = decimal.Parse(worksheet.Cells[row, 11].GetValue<string>());
                var buyUsdTotalCost = quantity * buyUsdPricePerStock;
                var buyRate = await GetBnbRate(buyDate);
                var buyBgnTotalCost = buyUsdTotalCost * buyRate;

                var sellDate = DateTime.Parse(worksheet.Cells[row, 12].GetValue<string>(), culture);
                var sellUsdPricePerStock = sellPrices[row - 3];
                var sellUsdTotalCost = quantity * sellUsdPricePerStock;
                var sellRate = await GetBnbRate(sellDate);
                var sellBgnTotalCost = sellUsdTotalCost * sellRate;

                var diff = Math.Round(sellBgnTotalCost - buyBgnTotalCost, 2, MidpointRounding.ToEven);
                bgnBenefit += diff;

                var sb = new StringBuilder();

                var printValues = new List<string>
                {
                    (row - 2).ToString(),
                    sellDate.ToString("dd.MM.yyyy"),
                    quantity.ToString(),
                    $"${sellUsdPricePerStock:.00}",
                    $"${sellUsdTotalCost:.00}",
                    sellRate.ToString(CultureInfo.InvariantCulture),
                    $"{sellBgnTotalCost:.00}",
                    buyDate.ToString("dd.MM.yyyy"),
                    $"${buyUsdPricePerStock:.00}",
                    $"${buyUsdTotalCost:.00}",
                    buyRate.ToString(CultureInfo.InvariantCulture),
                    $"{buyBgnTotalCost:.00}",
                    diff.ToString(CultureInfo.InvariantCulture)
                };

                printValues.ForEach(x => sb.AppendFormat("{0}\t", x));
                Console.WriteLine(sb);
            }

            var tax = Math.Round(bgnBenefit * 0.1M, 2, MidpointRounding.ToEven);
            Console.WriteLine($"Total benefit: {bgnBenefit:.00} лв. -> {tax} лв.");
        }

        /// <summary>
        /// Goes to bnb.bg and tries to retrieve the last 5 rates starting from the given date.
        /// E.g. 2020-10-16 => we'll ask for [2020-10-11..2020-10-16].
        /// From the provided response (csv) we'll take the last entry. This represents either the
        /// date that was requested or the last rate before that date.
        /// </summary>
        /// <param name="date">for which date to retrieve the USD rate</param>
        /// <returns>the rate for the provided date or the rate for the previous (working) day</returns>
        private static async Task<decimal> GetBnbRate(DateTime date)
        {
            var queryParams = HttpUtility.ParseQueryString(string.Empty);

            queryParams["downloadOper"] = "true";
            queryParams["group1"] = "second";
            queryParams["valutes"] = "USD";
            queryParams["search"] = "true";
            queryParams["showChart"] = "false";
            queryParams["showChartButton"] = "true";
            queryParams["type"] = "CSV";

            var from = date.AddDays(-5).Date;

            queryParams["periodStartDays"] = from.Day.ToString();
            queryParams["periodStartMonths"] = from.Month.ToString();
            queryParams["periodStartYear"] = from.Year.ToString();
            queryParams["periodEndDays"] = date.Day.ToString();
            queryParams["periodEndMonths"] = date.Month.ToString();
            queryParams["periodEndYear"] = date.Year.ToString();

            var uriBuilder = new UriBuilder(
            "https", "www.bnb.bg", 443, "Statistics/StExternalSector/StExchangeRates/StERForeignCurrencies/index.htm");
            uriBuilder.Query = queryParams.ToString();

            var httpResponse = await HttpClient.GetAsync(uriBuilder.ToString());
            var csv = await httpResponse.Content.ReadAsStringAsync();

            var lines = csv.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            var lastLine = lines.Last();

            var values = lastLine.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            var value = values.Last();

            return decimal.Parse(value);
        }
    }
}