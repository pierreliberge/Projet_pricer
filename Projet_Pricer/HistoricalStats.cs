using System;
using System.Collections.Generic;
using System.Linq;

namespace Projet_Pricer
{
    public static class HistoricalStats
    {
        /// <summary>
        /// Calcule les rendements log journaliers pour chaque ticker.
        /// Input : prices[date][ticker] = prix
        /// Output : returns[ticker] = liste des log-returns ln(S_t / S_{t-1})
        /// </summary>
        public static Dictionary<string, List<double>> LogReturns(
            SortedDictionary<DateTime, Dictionary<string, double>> prices,
            List<string> tickers)
        {
            var rets = new Dictionary<string, List<double>>();
            foreach (var t in tickers)
                rets[t] = new List<double>();

            Dictionary<string, double> prev = null;

            foreach (var kv in prices)
            {
                var cur = kv.Value;

                if (prev != null)
                {
                    // On ne calcule les returns que si tous les tickers existent aux deux dates
                    bool ok = true;
                    for (int i = 0; i < tickers.Count; i++)
                    {
                        string t = tickers[i];
                        if (!prev.ContainsKey(t) || !cur.ContainsKey(t))
                        {
                            ok = false;
                            break;
                        }
                    }

                    if (ok)
                    {
                        for (int i = 0; i < tickers.Count; i++)
                        {
                            string t = tickers[i];
                            double s0 = prev[t];
                            double s1 = cur[t];

                            // sécurité
                            if (s0 > 0.0 && s1 > 0.0)
                                rets[t].Add(Math.Log(s1 / s0));
                        }
                    }
                }

                prev = cur;
            }

            return rets;
        }

        /// <summary>
        /// Volatilité annualisée à partir de log-returns journaliers.
        /// sigma = std(returns) * sqrt(252)
        /// </summary>
        public static double AnnualizedVol(List<double> logReturns, int tradingDays = 252)
        {
            if (logReturns == null || logReturns.Count < 2)
                throw new ArgumentException("Not enough returns to compute vol.");

            double mean = logReturns.Average();

            double sumSq = 0.0;
            for (int i = 0; i < logReturns.Count; i++)
            {
                double x = logReturns[i] - mean;
                sumSq += x * x;
            }

            // variance "population" (division par N) -> suffisant pour ce projet
            double var = sumSq / logReturns.Count;
            double stdev = Math.Sqrt(Math.Max(var, 1e-18));

            return stdev * Math.Sqrt(tradingDays);
        }

        /// <summary>
        /// Matrice de corrélation NxN des log-returns.
        /// Les séries sont tronquées à la même longueur (minLen) pour être comparables.
        /// </summary>
        public static double[,] CorrelationMatrix(
            Dictionary<string, List<double>> returnsByTicker,
            List<string> tickers)
        {
            if (tickers == null || tickers.Count == 0)
                throw new ArgumentException("Tickers required.");

            int n = tickers.Count;

            // On s'aligne sur la longueur min pour éviter les décalages
            int minLen = int.MaxValue;
            for (int i = 0; i < n; i++)
            {
                string t = tickers[i];
                if (!returnsByTicker.ContainsKey(t))
                    throw new ArgumentException("Missing returns for " + t);

                minLen = Math.Min(minLen, returnsByTicker[t].Count);
            }

            if (minLen < 2)
                throw new ArgumentException("Not enough aligned returns to compute correlation.");

            // Centrage + std
            var centered = new double[n][];
            var std = new double[n];

            for (int i = 0; i < n; i++)
            {
                string t = tickers[i];
                var r = returnsByTicker[t].Take(minLen).ToArray();

                double mean = r.Average();
                centered[i] = new double[minLen];

                double sumSq = 0.0;
                for (int k = 0; k < minLen; k++)
                {
                    double x = r[k] - mean;
                    centered[i][k] = x;
                    sumSq += x * x;
                }

                double var = sumSq / minLen;
                std[i] = Math.Sqrt(Math.Max(var, 1e-18));
            }

            // Corr = cov / (std_i * std_j)
            var corr = new double[n, n];

            for (int i = 0; i < n; i++)
            {
                corr[i, i] = 1.0;

                for (int j = i + 1; j < n; j++)
                {
                    double cov = 0.0;
                    for (int k = 0; k < minLen; k++)
                        cov += centered[i][k] * centered[j][k];
                    cov /= minLen;

                    double c = cov / (std[i] * std[j]);
                    corr[i, j] = c;
                    corr[j, i] = c;
                }
            }

            return corr;
        }
    }
}
