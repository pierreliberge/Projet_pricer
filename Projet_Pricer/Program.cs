using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Projet_Pricer
{
    internal class Program
    {
        // =========================================================
        // Utilitaire: garder les N dernières observations (dates)
        //  - 126 ~ 6 mois de bourse
        //  - 252 ~ 1 an de bourse
        // =========================================================
        static SortedDictionary<DateTime, Dictionary<string, double>> TakeLastNObs(
            SortedDictionary<DateTime, Dictionary<string, double>> prices, int n)
        {
            var last = prices.Skip(Math.Max(0, prices.Count - n));
            return new SortedDictionary<DateTime, Dictionary<string, double>>(
                last.ToDictionary(kv => kv.Key, kv => kv.Value)
            );
        }

        static void Main(string[] args)
        {
            // =========================
            // 1) LOAD DATA CAC40 (Excel)
            // =========================
            var prices = MarketDataLoader.LoadWidePricesFromExcel(System.IO.Path.Combine("data", "data_CAC40.xlsx"), "Feuil1");

            Console.WriteLine("=== Chargement Excel OK ===");
            Console.WriteLine($"Nb dates           = {prices.Count}");
            Console.WriteLine($"Première date      = {prices.First().Key:dd/MM/yyyy}");
            Console.WriteLine($"Dernière date      = {prices.Last().Key:dd/MM/yyyy}");
            Console.WriteLine($"Nb tickers (col)   = {prices.First().Value.Count}");

            // Date de pricing = dernière date dispo (ex: 26/01/2026)
            DateTime asOf = prices.Last().Key;

            // Tous les tickers disponibles (ordre stable)
            var allTickers = prices.First().Value.Keys.OrderBy(x => x).ToList();

            Console.WriteLine("\n=== Tickers disponibles ===");
            for (int i = 0; i < allTickers.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {allTickers[i]}");
            }

            // =========================
            // 2) SÉLECTION DES TICKERS POUR LE PANIER
            // =========================
            List<string> selectedTickers = new List<string>();

            Console.WriteLine("\n=== Sélection des tickers pour le panier ===");
            Console.WriteLine("Choisissez entre 2 et " + allTickers.Count + " tickers.");
            Console.WriteLine("Entrez les numéros séparés par des virgules (ex: 1,3,5,7) ou 'ALL' pour tous:");

            while (true)
            {
                Console.Write("> ");
                string input = Console.ReadLine()?.Trim().ToUpper();

                if (input == "ALL")
                {
                    selectedTickers = new List<string>(allTickers);
                    break;
                }

                try
                {
                    var indices = input.Split(',')
                        .Select(s => int.Parse(s.Trim()))
                        .Where(idx => idx >= 1 && idx <= allTickers.Count)
                        .Distinct()
                        .OrderBy(x => x)
                        .ToList();

                    if (indices.Count < 2)
                    {
                        Console.WriteLine("❌ Vous devez sélectionner au moins 2 tickers !");
                        continue;
                    }

                    selectedTickers = indices.Select(idx => allTickers[idx - 1]).ToList();
                    break;
                }
                catch
                {
                    Console.WriteLine("❌ Format invalide. Réessayez (ex: 1,3,5,7 ou ALL)");
                }
            }

            Console.WriteLine("\n✅ Tickers sélectionnés pour le panier:");
            foreach (var t in selectedTickers)
                Console.WriteLine($"  - {t}");

            // Filtrer les données de prix pour ne garder que les tickers sélectionnés
            var filteredPrices = new SortedDictionary<DateTime, Dictionary<string, double>>();
            foreach (var dateKv in prices)
            {
                var filteredMap = new Dictionary<string, double>();
                bool allPresent = true;

                foreach (var ticker in selectedTickers)
                {
                    if (dateKv.Value.ContainsKey(ticker))
                        filteredMap[ticker] = dateKv.Value[ticker];
                    else
                    {
                        allPresent = false;
                        break;
                    }
                }

                if (allPresent)
                    filteredPrices[dateKv.Key] = filteredMap;
            }

            prices = filteredPrices;
            var tickers = selectedTickers;

            Console.WriteLine($"\n📊 Nb dates après filtrage = {prices.Count}");

            // =========================
            // 3) Returns / vols / corr (FULL SAMPLE)
            // =========================
            var rets = HistoricalStats.LogReturns(prices, tickers);

            Console.WriteLine("\n=== Volatilités historiques (échantillon complet) ===");
            foreach (var t in tickers)
                Console.WriteLine($"{t,-15} vol hist = {HistoricalStats.AnnualizedVol(rets[t]):P2}");

            // --- Demande du dividende q ---
            double qUser;
            while (true)
            {
                Console.Write("\nEntrez le dividende q à appliquer à tous les actifs (ex: 0.02) : ");
                string input = Console.ReadLine();

                if (double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out qUser) ||
                    double.TryParse(input, NumberStyles.Float, CultureInfo.GetCultureInfo("fr-FR"), out qUser))
                    break;

                Console.WriteLine("❌ Valeur invalide, veuillez entrer un nombre valide !");
            }

            // --- Création des assets ---
            var assets = new List<Asset>();
            foreach (var t in tickers)
            {
                double spot = prices.Last().Value[t];
                double vol = HistoricalStats.AnnualizedVol(rets[t]);
                assets.Add(new Asset(t, spot, q: qUser, volConst: vol));
            }

            // Basket équipondéré
            double[] w = Enumerable.Repeat(1.0 / tickers.Count, tickers.Count).ToArray();
            var basket = new Basket(assets, w);

            // Matrice de corrélation (sur log-returns, full)
            double[,] corr = HistoricalStats.CorrelationMatrix(rets, tickers);

            Console.WriteLine("\n=== Matrice de corrélation (aperçu) ===");
            int maxDisplay = Math.Min(5, tickers.Count);
            for (int i = 0; i < maxDisplay; i++)
            {
                for (int j = 0; j < maxDisplay; j++)
                {
                    Console.Write($"{corr[i, j],7:F3} ");
                }
                Console.WriteLine();
            }
            if (tickers.Count > maxDisplay)
                Console.WriteLine("... (matrice tronquée)");

            // =========================
            // 4) Option (exemple)
            // =========================
            double basketSpot = 0.0;
            for (int i = 0; i < basket.Dim; i++)
                basketSpot += basket.Weights[i] * basket.Assets[i].Spot;

            Console.Write("\nEntrez la maturité de l'option (en années, ex: 1, 2, 3) : ");
            double T;
            while (!double.TryParse(Console.ReadLine(), NumberStyles.Float, CultureInfo.InvariantCulture, out T) &&
                   !double.TryParse(Console.ReadLine(), NumberStyles.Float, CultureInfo.GetCultureInfo("fr-FR"), out T))
            {
                Console.Write("❌ Valeur invalide. Entrez une maturité valide : ");
            }

            var opt = new OptionSpec(OptionType.Call, strike: basketSpot, maturity: T);

            // =========================
            // 5) Modèles H1 / H2
            // =========================
            // H1 : taux constant
            double r = 0.03;
            var modelH1 = new MarketModelH1(basket, r, corr);

            // ====== H1 recalibré sur fenêtres 6M et 1Y (vols + corr) ======
            var prices6m = TakeLastNObs(prices, 126);
            var prices1y = TakeLastNObs(prices, 252);

            // --- 6M ---
            var rets6m = HistoricalStats.LogReturns(prices6m, tickers);
            var corr6m = HistoricalStats.CorrelationMatrix(rets6m, tickers);

            var assets6m = new List<Asset>();
            foreach (var t in tickers)
            {
                double spot = prices.Last().Value[t];
                double vol6m = HistoricalStats.AnnualizedVol(rets6m[t]);
                assets6m.Add(new Asset(t, spot, q: qUser, volConst: vol6m));
            }
            var basket6m = new Basket(assets6m, w);
            var modelH1_6m = new MarketModelH1(basket6m, r, corr6m);

            // --- 1Y ---
            var rets1y = HistoricalStats.LogReturns(prices1y, tickers);
            var corr1y = HistoricalStats.CorrelationMatrix(rets1y, tickers);

            var assets1y = new List<Asset>();
            foreach (var t in tickers)
            {
                double spot = prices.Last().Value[t];
                double vol1y = HistoricalStats.AnnualizedVol(rets1y[t]);
                assets1y.Add(new Asset(t, spot, q: qUser, volConst: vol1y));
            }
            var basket1y = new Basket(assets1y, w);
            var modelH1_1y = new MarketModelH1(basket1y, r, corr1y);

            // Diagnostics vols moyennes
            Console.WriteLine("\n=== Calibration H1 : fenêtres de vol/corr ===");
            Console.WriteLine($"Full sample: nb dates={prices.Count}  | avg vol={basket.Assets.Average(a => a.VolConst):P2}");
            Console.WriteLine($"6M window  : nb dates={prices6m.Count} | avg vol={basket6m.Assets.Average(a => a.VolConst):P2}");
            Console.WriteLine($"1Y window  : nb dates={prices1y.Count} | avg vol={basket1y.Assets.Average(a => a.VolConst):P2}");

            // H2 : courbe de taux + vols déterministes
            var rateCurve = MarketDataLoader.LoadRateCurveFromExcel(
                Path.Combine("data", "vol_impli.xlsx"),
                "TAUX"
            );

            // --- VOLS implicites ATM depuis vol_impli.xlsx ---
            Dictionary<string, VolCurve> volCurves;
            try
            {
                var allVolCurves = MarketDataLoader.LoadAtmVolCurvesFromExcel(
                    xlsxPath: System.IO.Path.Combine("data", "vol_impli.xlsx"),
                    asOfDate: asOf,
                    sheetName: "VOL_LONG",
                    tickerCol: "Ticker",
                    expiryCol: "Expiry",
                    volPctCol: "ATMVolPct",
                    dayCountBasis: 365.0
                );

                // Filtrer pour ne garder que les tickers sélectionnés
                volCurves = new Dictionary<string, VolCurve>();
                foreach (var t in tickers)
                {
                    if (!allVolCurves.ContainsKey(t))
                        throw new Exception($"VOL_LONG: missing vol curve for ticker '{t}'");
                    volCurves[t] = allVolCurves[t];
                }

                Console.WriteLine($"\n=== Vol implicites chargées (vol_impli.xlsx / VOL_LONG) ===");
                Console.WriteLine($"AsOf = {asOf:dd/MM/yyyy} | Nb tickers vol = {volCurves.Count}");

                // =========================
                // CHECK : vol hist vs vol implicite équivalente (T = 1Y)
                // =========================
                Console.WriteLine("\n=== Check vol hist vs vol implicite eq (T = 1Y) ===");

                double Tcheck = 1.0;
                var rows = new List<(string tkr, double volHist, double volImplEq)>();

                foreach (var a in basket.Assets)
                {
                    double volHist = a.VolConst;

                    var vc = volCurves[a.Ticker];
                    double intSig2 = vc.IntegralVol2(Tcheck);
                    double volImplEq = Math.Sqrt(Math.Max(intSig2 / Tcheck, 0.0));

                    rows.Add((a.Ticker, volHist, volImplEq));
                }

                foreach (var row in rows.OrderByDescending(x => x.volImplEq - x.volHist))
                {
                    Console.WriteLine(
                        $"{row.tkr,-15} | vol hist = {row.volHist:P2} | vol impl eq(1Y) = {row.volImplEq:P2} | diff = {(row.volImplEq - row.volHist):P2}"
                    );
                }

                double avgHist = rows.Average(x => x.volHist);
                double avgImpl = rows.Average(x => x.volImplEq);

                Console.WriteLine("\n--- Moyenne équipondérée panier ---");
                Console.WriteLine($"Hist = {avgHist:P2}");
                Console.WriteLine($"Impl eq(1Y) = {avgImpl:P2}");
                Console.WriteLine($"Diff = {(avgImpl - avgHist):P2}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n[WARN] Impossible de charger vol_impli.xlsx / VOL_LONG (" + ex.Message + ")");
                Console.WriteLine("[WARN] Fallback: vol-curves plates = vols historiques");

                volCurves = new Dictionary<string, VolCurve>();
                foreach (var a in basket.Assets)
                {
                    double s = a.VolConst;
                    volCurves[a.Ticker] = new VolCurve(new List<(double t, double vol)>
                    {
                        (0.0833333333, s),
                        (0.25,         s),
                        (0.50,         s),
                        (1.00,         s),
                        (2.00,         s)
                    });
                }
            }

            var modelH2 = new MarketModelH2(basket, rateCurve, volCurves, corr);

            // =========================
            // 6) Snapshot prix H1 (Full/6M/1Y) vs H2 (MM)
            // =========================
            Console.WriteLine("\n=== Snapshot (MM) : H1 Full vs H1 6M vs H1 1Y vs H2 ===");
            Console.WriteLine($"H1 Full (MM) = {BasketPricers.MomentMatchingPrice(modelH1, opt):F6}");
            Console.WriteLine($"H1 6M   (MM) = {BasketPricers.MomentMatchingPrice(modelH1_6m, opt):F6}");
            Console.WriteLine($"H1 1Y   (MM) = {BasketPricers.MomentMatchingPrice(modelH1_1y, opt):F6}");
            Console.WriteLine($"H2      (MM) = {BasketPricers.MomentMatchingPrice(modelH2, opt):F6}");

            // =========================
            // 7) Menu principal
            // =========================
            while (true)
            {
                Console.WriteLine("\n=== Basket Option Pricer ===");
                Console.WriteLine($"AsOf {asOf:dd/MM/yyyy} | Basket spot ~ {basketSpot:F4} | Strike={opt.Strike:F4} | T={opt.Maturity:F2}y");
                Console.WriteLine($"Panier: {tickers.Count} tickers équipondérés");
                Console.WriteLine("1) Moment Matching (H1 Full)");
                Console.WriteLine("2) Monte Carlo (H1 Full) + control variate (geo)");
                Console.WriteLine("3) Moment Matching (H2)");
                Console.WriteLine("4) Monte Carlo (H2) + control variate (geo)");
                Console.WriteLine("5) Moment Matching (H1 6M window)");
                Console.WriteLine("6) Moment Matching (H1 1Y window)");
                Console.WriteLine("0) Quit");
                Console.Write("Choix: ");
                var choice = Console.ReadLine();

                if (choice == "0") return;

                try
                {
                    switch (choice)
                    {
                        case "1":
                            {
                                var price = BasketPricers.MomentMatchingPrice(modelH1, opt);
                                Console.WriteLine($"✅ MM H1 Full price = {price:F6}");
                                break;
                            }
                        case "2":
                            {
                                Console.Write("Nb simulations (ex 200000): ");
                                int n = int.Parse(Console.ReadLine() ?? "200000");
                                Console.Write("Seed (ex 42): ");
                                int seed = int.Parse(Console.ReadLine() ?? "42");

                                var res = BasketPricers.MonteCarloPriceWithControlVariate(modelH1, opt, n, seed);
                                Console.WriteLine($"✅ MC H1 Full price (CV geo) = {res.Price:F6}");
                                Console.WriteLine($"   Variance(estimator)       = {res.Variance:F10}");
                                Console.WriteLine($"   StdError                  = {Math.Sqrt(res.Variance):F6}");
                                Console.WriteLine($"   IC 95% approx             = [{res.Price - 1.96 * Math.Sqrt(res.Variance):F6} ; {res.Price + 1.96 * Math.Sqrt(res.Variance):F6}]");
                                break;
                            }
                        case "3":
                            {
                                var price = BasketPricers.MomentMatchingPrice(modelH2, opt);
                                Console.WriteLine($"✅ MM H2 price = {price:F6}");
                                break;
                            }
                        case "4":
                            {
                                Console.Write("Nb simulations (ex 200000): ");
                                int n = int.Parse(Console.ReadLine() ?? "200000");
                                Console.Write("Seed (ex 42): ");
                                int seed = int.Parse(Console.ReadLine() ?? "42");

                                var res = BasketPricers.MonteCarloPriceWithControlVariate(modelH2, opt, n, seed);
                                Console.WriteLine($"✅ MC H2 price (CV geo) = {res.Price:F6}");
                                Console.WriteLine($"   Variance(estimator)  = {res.Variance:F10}");
                                Console.WriteLine($"   StdError             = {Math.Sqrt(res.Variance):F6}");
                                Console.WriteLine($"   IC 95% approx        = [{res.Price - 1.96 * Math.Sqrt(res.Variance):F6} ; {res.Price + 1.96 * Math.Sqrt(res.Variance):F6}]");
                                break;
                            }
                        case "5":
                            {
                                var price = BasketPricers.MomentMatchingPrice(modelH1_6m, opt);
                                Console.WriteLine($"✅ MM H1 6M price = {price:F6}");
                                break;
                            }
                        case "6":
                            {
                                var price = BasketPricers.MomentMatchingPrice(modelH1_1y, opt);
                                Console.WriteLine($"✅ MM H1 1Y price = {price:F6}");
                                break;
                            }
                        default:
                            Console.WriteLine("❌ Choix invalide.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("❌ Erreur: " + ex.Message);
                }
            }
        }
    }
}