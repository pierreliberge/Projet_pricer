using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace Projet_Pricer
{
    public static class MarketDataLoader
    {
        // ============================================================
        // 1) PRICES: Excel "wide" (Dates + colonnes tickers)
        // Retourne: Date -> (Ticker -> Prix)
        // ============================================================
        public static SortedDictionary<DateTime, Dictionary<string, double>> LoadWidePricesFromExcel(
            string xlsxPath,
            string sheetName = "Feuil1")
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                });

                var table = ds.Tables[sheetName]
                    ?? throw new ArgumentException("Sheet '" + sheetName + "' not found.");

                // tickers = toutes les colonnes sauf "Dates"
                var tickers = new List<string>();
                for (int c = 1; c < table.Columns.Count; c++)
                    tickers.Add(table.Columns[c].ColumnName.Trim());

                var data = new SortedDictionary<DateTime, Dictionary<string, double>>();

                foreach (DataRow row in table.Rows)
                {
                    if (row[0] == DBNull.Value) continue;

                    DateTime date;
                    if (row[0] is DateTime)
                        date = (DateTime)row[0];
                    else if (!DateTime.TryParse(row[0].ToString(), out date))
                        continue;

                    var map = new Dictionary<string, double>();

                    for (int i = 0; i < tickers.Count; i++)
                    {
                        var cell = row[i + 1];
                        if (cell == DBNull.Value) break;

                        double px;
                        var s = cell.ToString();

                        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out px) && px > 0)
                            map[tickers[i]] = px;
                        else if (double.TryParse(s, NumberStyles.Float, CultureInfo.GetCultureInfo("fr-FR"), out px) && px > 0)
                            map[tickers[i]] = px;
                    }

                    // on garde uniquement les dates où on a tous les tickers
                    if (map.Count == tickers.Count)
                        data[date] = map;
                }

                return data;
            }
        }

        // ============================================================
        // 2) VOLS: Excel "long" (Ticker | Expiry | ATMVolPct)
        // Retourne: Ticker -> VolCurve (vol en décimal, ex 0.2345)
        // ============================================================
        public static Dictionary<string, VolCurve> LoadAtmVolCurvesFromExcel(
            string xlsxPath,
            DateTime asOfDate,
            string sheetName = "VOL_LONG",
            string tickerCol = "Ticker",
            string expiryCol = "Expiry",
            string volPctCol = "ATMVolPct",
            double dayCountBasis = 365.0)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                });

                var table = ds.Tables[sheetName];
                if (table == null)
                    throw new ArgumentException("Sheet '" + sheetName + "' not found in '" + xlsxPath + "'.");

                // Vérif colonnes
                if (!table.Columns.Contains(tickerCol) ||
                    !table.Columns.Contains(expiryCol) ||
                    !table.Columns.Contains(volPctCol))
                {
                    throw new ArgumentException(
                        "Sheet '" + sheetName + "' must contain columns: '" + tickerCol + "', '" + expiryCol + "', '" + volPctCol + "'.");
                }

                var pointsByTicker = new Dictionary<string, List<(double t, double vol)>>();

                foreach (DataRow row in table.Rows)
                {
                    if (row[tickerCol] == DBNull.Value || row[expiryCol] == DBNull.Value || row[volPctCol] == DBNull.Value)
                        continue;

                    var tickerObj = row[tickerCol];
                    var expiryObj = row[expiryCol];
                    var volObj = row[volPctCol];

                    if (tickerObj == null || expiryObj == null || volObj == null) continue;

                    string ticker = tickerObj.ToString().Trim();
                    if (string.IsNullOrWhiteSpace(ticker)) continue;

                    // --- Expiry ---
                    DateTime expiry;
                    if (expiryObj is DateTime)
                    {
                        expiry = (DateTime)expiryObj;
                    }
                    else
                    {
                        string s = expiryObj.ToString().Trim();

                        // Invariant formats
                        if (!DateTime.TryParseExact(
                                s,
                                new[] { "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d MMM yyyy", "d MMMM yyyy" },
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                                out expiry))
                        {
                            // fallback FR
                            if (!DateTime.TryParse(s, CultureInfo.GetCultureInfo("fr-FR"), DateTimeStyles.None, out expiry))
                                continue;
                        }
                    }

                    double days = (expiry.Date - asOfDate.Date).TotalDays;
                    if (days <= 0) continue;

                    double t = days / dayCountBasis;

                    // --- Vol % ---
                    string volStr = volObj.ToString().Replace("%", "").Trim();

                    double volPct;
                    if (!double.TryParse(volStr, NumberStyles.Float, CultureInfo.InvariantCulture, out volPct))
                    {
                        if (!double.TryParse(volStr, NumberStyles.Float, CultureInfo.GetCultureInfo("fr-FR"), out volPct))
                            continue;
                    }

                    double vol = volPct / 100.0; // décimal

                    List<(double t, double vol)> list;
                    if (!pointsByTicker.TryGetValue(ticker, out list))
                    {
                        list = new List<(double t, double vol)>();
                        pointsByTicker[ticker] = list;
                    }
                    list.Add((t, vol));
                }

                // Construire VolCurve
                var res = new Dictionary<string, VolCurve>();

                foreach (var kv in pointsByTicker)
                {
                    var ticker = kv.Key;

                    var pts = kv.Value
                        .OrderBy(p => p.t)
                        .GroupBy(p => p.t)
                        .Select(g => g.Last())
                        .ToList();

                    if (pts.Count < 2)
                        throw new Exception("Not enough vol points for '" + ticker + "' (need >=2).");

                    res[ticker] = new VolCurve(pts);
                }

                return res;
            }
        }
        // ============================================================
        // 3) TAUX: Excel "TAUX" (Tenor | Yield)
        // Tenor: 3M, 6M, 1Y, ...
        // Yield: en %, format FR (ex 2,213)
        // Retourne: RateCurve (t en années, z en décimal)
        // ============================================================
        public static RateCurve LoadRateCurveFromExcel(
            string xlsxPath,
            string sheetName = "TAUX",
            string tenorCol = "Tenor",
            string yieldCol = "Yield")
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                });

                var table = ds.Tables[sheetName]
                    ?? throw new ArgumentException($"Sheet '{sheetName}' not found.");

                if (!table.Columns.Contains(tenorCol) || !table.Columns.Contains(yieldCol))
                    throw new ArgumentException(
                        $"Sheet '{sheetName}' must contain columns '{tenorCol}' and '{yieldCol}'.");

                var points = new List<(double t, double z)>();

                foreach (DataRow row in table.Rows)
                {
                    if (row[tenorCol] == DBNull.Value || row[yieldCol] == DBNull.Value)
                        continue;

                    // ---------- Tenor ----------
                    string tenorStr = row[tenorCol].ToString().Trim().ToUpper();
                    if (string.IsNullOrWhiteSpace(tenorStr))
                        continue;

                    double t;
                    if (tenorStr.EndsWith("M"))
                    {
                        if (!double.TryParse(tenorStr.Replace("M", ""),
                                NumberStyles.Float,
                                CultureInfo.InvariantCulture,
                                out double months))
                            continue;

                        t = months / 12.0;
                    }
                    else if (tenorStr.EndsWith("Y"))
                    {
                        if (!double.TryParse(tenorStr.Replace("Y", ""),
                                NumberStyles.Float,
                                CultureInfo.InvariantCulture,
                                out double years))
                            continue;

                        t = years;
                    }
                    else
                    {
                        continue; // tenor inconnu
                    }

                    if (t <= 0) continue;

                    // ---------- Yield (%) ----------
                    string yieldStr = row[yieldCol].ToString().Trim();

                    if (!double.TryParse(yieldStr,
                            NumberStyles.Float,
                            CultureInfo.GetCultureInfo("fr-FR"),
                            out double yieldPct))
                    {
                        if (!double.TryParse(yieldStr,
                                NumberStyles.Float,
                                CultureInfo.InvariantCulture,
                                out yieldPct))
                            continue;
                    }

                    double z = yieldPct / 100.0; // % -> décimal

                    points.Add((t, z));
                }

                if (points.Count < 2)
                    throw new Exception("Not enough rate points (need >= 2).");

                return RateCurve.FromZeroRates(points.OrderBy(p => p.t).ToList());
            }
        }

    }
}
