using System;
using System.Collections.Generic;
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
            Console.WriteLine("Tickers: " + string.Join(", ", prices.First().Value.Keys));

            // Date de pricing = dernière date dispo (ex: 26/01/2026)
            DateTime asOf = prices.Last().Key;

            // Tickers (ordre stable)
            var tickers = prices.First().Value.Keys.OrderBy(x => x).ToList();

            // =========================
            // 2) Returns / vols / corr (FULL SAMPLE)
            // =========================
            var rets = HistoricalStats.LogReturns(prices, tickers);

            foreach (var t in tickers.Take(3))
                Console.WriteLine($"{t} vol hist = {HistoricalStats.AnnualizedVol(rets[t]):P2}");

            // Assets : spot = dernier PX_LAST, q=0, volConst = vol hist annualisée (sert H1 full)
            var assets = new List<Asset>();
            foreach (var t in tickers)
            {
                double spot = prices.Last().Value[t];
                double vol = HistoricalStats.AnnualizedVol(rets[t]);
                assets.Add(new Asset(t, spot, q: 0.0, volConst: vol));
            }

            // Basket équipondéré
            double[] w = Enumerable.Repeat(1.0 / tickers.Count, tickers.Count).ToArray();
            var basket = new Basket(assets, w);

            // Matrice de corrélation (sur log-returns, full)
            double[,] corr = HistoricalStats.CorrelationMatrix(rets, tickers);

            Console.WriteLine("\n=== Check corr matrix (quelques valeurs) ===");
            Console.WriteLine($"corr(AIR, BN) = {corr[0, 1]:F3}");
            Console.WriteLine($"corr(AIR, CAP)= {corr[0, 2]:F3}");
            Console.WriteLine($"corr(BN, CAP) = {corr[1, 2]:F3}");
            Console.WriteLine($"corr(MC, OR)  = {corr[7, 8]:F3}");

            // =========================
            // 3) Option (exemple)
            // =========================
            double basketSpot = 0.0;
            for (int i = 0; i < basket.Dim; i++)
                basketSpot += basket.Weights[i] * basket.Assets[i].Spot;

            var opt = new OptionSpec(OptionType.Call, strike: basketSpot, maturity: 1.0);

            // =========================
            // 4) Modèles H1 / H2
            // =========================
            // H1 : taux constant (tu remplaceras plus tard par ta courbe)
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
                double spot = prices.Last().Value[t]; // spot asOf (inchangé)
                double vol6m = HistoricalStats.AnnualizedVol(rets6m[t]);
                assets6m.Add(new Asset(t, spot, q: 0.0, volConst: vol6m));
            }
            var basket6m = new Basket(assets6m, w);
            var modelH1_6m = new MarketModelH1(basket6m, r, corr6m);

            // --- 1Y ---
            var rets1y = HistoricalStats.LogReturns(prices1y, tickers);
            var corr1y = HistoricalStats.CorrelationMatrix(rets1y, tickers);

            var assets1y = new List<Asset>();
            foreach (var t in tickers)
            {
                double spot = prices.Last().Value[t]; // spot asOf (inchangé)
                double vol1y = HistoricalStats.AnnualizedVol(rets1y[t]);
                assets1y.Add(new Asset(t, spot, q: 0.0, volConst: vol1y));
            }
            var basket1y = new Basket(assets1y, w);
            var modelH1_1y = new MarketModelH1(basket1y, r, corr1y);

            // Diagnostics vols moyennes
            Console.WriteLine("\n=== Calibration H1 : fenêtres de vol/corr ===");
            Console.WriteLine($"Full sample: nb dates={prices.Count}  | avg vol={basket.Assets.Average(a => a.VolConst):P2}");
            Console.WriteLine($"6M window  : nb dates={prices6m.Count} | avg vol={basket6m.Assets.Average(a => a.VolConst):P2}");
            Console.WriteLine($"1Y window  : nb dates={prices1y.Count} | avg vol={basket1y.Assets.Average(a => a.VolConst):P2}");

            // H2 : courbe de taux + vols déterministes
            var rateCurve = RateCurve.FromZeroRates(new List<(double t, double z)>
            {
                (0.0833333333, 0.030),
                (0.25,         0.030),
                (0.50,         0.030),
                (1.00,         0.030),
                (2.00,         0.030)
            });

            // --- VOLS implicites ATM depuis vol_impli.xlsx ---
            Dictionary<string, VolCurve> volCurves;
            try
            {
                volCurves = MarketDataLoader.LoadAtmVolCurvesFromExcel(
                    xlsxPath: System.IO.Path.Combine("data", "vol_impli.xlsx"),
                    asOfDate: asOf,
                    sheetName: "VOL_LONG",
                    tickerCol: "Ticker",
                    expiryCol: "Expiry",
                    volPctCol: "ATMVolPct",
                    dayCountBasis: 365.0
                );

                foreach (var a in basket.Assets)
                {
                    if (!volCurves.ContainsKey(a.Ticker))
                        throw new Exception("VOL_LONG: missing vol curve for ticker '" + a.Ticker + "'");
                }

                Console.WriteLine($"\n=== Vol implicites chargées (vol_impli.xlsx / VOL_LONG) ===");
                Console.WriteLine($"AsOf = {asOf:dd/MM/yyyy} | Nb tickers vol = {volCurves.Count}");

                // =========================
                // CHECK #1 : vol hist vs vol implicite équivalente (T = 1Y)
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
            // 4bis) Snapshot prix H1 (Full/6M/1Y) vs H2 (MM)
            // =========================
            Console.WriteLine("\n=== Snapshot (MM) : H1 Full vs H1 6M vs H1 1Y vs H2 ===");
            Console.WriteLine($"H1 Full (MM) = {BasketPricers.MomentMatchingPrice(modelH1, opt):F6}");
            Console.WriteLine($"H1 6M   (MM) = {BasketPricers.MomentMatchingPrice(modelH1_6m, opt):F6}");
            Console.WriteLine($"H1 1Y   (MM) = {BasketPricers.MomentMatchingPrice(modelH1_1y, opt):F6}");
            Console.WriteLine($"H2      (MM) = {BasketPricers.MomentMatchingPrice(modelH2, opt):F6}");

            // =========================
            // 5) Menu
            // =========================
            while (true)
            {
                Console.WriteLine("\n=== Basket Option Pricer ===");
                Console.WriteLine($"AsOf {asOf:dd/MM/yyyy} | Basket spot ~ {basketSpot:F4} | Strike={opt.Strike:F4} | T={opt.Maturity:F2}y");
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
                                Console.WriteLine($"MM H1 Full price = {price:F6}");
                                break;
                            }
                        case "2":
                            {
                                Console.Write("Nb simulations (ex 200000): ");
                                int n = int.Parse(Console.ReadLine() ?? "200000");
                                Console.Write("Seed (ex 42): ");
                                int seed = int.Parse(Console.ReadLine() ?? "42");

                                var res = BasketPricers.MonteCarloPriceWithControlVariate(modelH1, opt, n, seed);
                                Console.WriteLine($"MC H1 Full price (CV geo) = {res.Price:F6}");
                                Console.WriteLine($"Variance(estimator)       = {res.Variance:F10}");
                                Console.WriteLine($"StdError                  = {Math.Sqrt(res.Variance):F6}");
                                Console.WriteLine($"IC 95% approx             = [{res.Price - 1.96 * Math.Sqrt(res.Variance):F6} ; {res.Price + 1.96 * Math.Sqrt(res.Variance):F6}]");
                                break;
                            }
                        case "3":
                            {
                                var price = BasketPricers.MomentMatchingPrice(modelH2, opt);
                                Console.WriteLine($"MM H2 price = {price:F6}");
                                break;
                            }
                        case "4":
                            {
                                Console.Write("Nb simulations (ex 200000): ");
                                int n = int.Parse(Console.ReadLine() ?? "200000");
                                Console.Write("Seed (ex 42): ");
                                int seed = int.Parse(Console.ReadLine() ?? "42");

                                var res = BasketPricers.MonteCarloPriceWithControlVariate(modelH2, opt, n, seed);
                                Console.WriteLine($"MC H2 price (CV geo) = {res.Price:F6}");
                                Console.WriteLine($"Variance(estimator)  = {res.Variance:F10}");
                                Console.WriteLine($"StdError             = {Math.Sqrt(res.Variance):F6}");
                                Console.WriteLine($"IC 95% approx        = [{res.Price - 1.96 * Math.Sqrt(res.Variance):F6} ; {res.Price + 1.96 * Math.Sqrt(res.Variance):F6}]");
                                break;
                            }
                        case "5":
                            {
                                var price = BasketPricers.MomentMatchingPrice(modelH1_6m, opt);
                                Console.WriteLine($"MM H1 6M price = {price:F6}");
                                break;
                            }
                        case "6":
                            {
                                var price = BasketPricers.MomentMatchingPrice(modelH1_1y, opt);
                                Console.WriteLine($"MM H1 1Y price = {price:F6}");
                                break;
                            }
                        default:
                            Console.WriteLine("Choix invalide.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Erreur: " + ex.Message);
                }
            }
        }
    }
}
