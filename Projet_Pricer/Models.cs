using System;
using System.Collections.Generic;
//Erwan
namespace Projet_Pricer
{
    public enum OptionType { Call, Put }

    public sealed class OptionSpec
    {
        public OptionType Type { get; }
        public double Strike { get; }
        public double Maturity { get; } // en années

        public OptionSpec(OptionType type, double strike, double maturity)
        {
            if (strike <= 0) throw new ArgumentException("Strike must be > 0");
            if (maturity <= 0) throw new ArgumentException("Maturity must be > 0");
            Type = type;
            Strike = strike;
            Maturity = maturity;
        }
    }

    public sealed class Asset
    {
        public string Ticker { get; }
        public double Spot { get; set; }
        public double DividendYield { get; set; } // q (continu)
        public double VolConst { get; set; }      // sigma constant (H1)

        public Asset(string ticker, double spot, double q, double volConst)
        {
            if (string.IsNullOrWhiteSpace(ticker)) throw new ArgumentException("Ticker required");
            if (spot <= 0) throw new ArgumentException("Spot must be > 0");
            if (volConst <= 0) throw new ArgumentException("Vol must be > 0");
            Ticker = ticker;
            Spot = spot;
            DividendYield = q;
            VolConst = volConst;
        }
    }

    public sealed class Basket
    {
        public List<Asset> Assets { get; }
        public double[] Weights { get; } // somme = 1 recommandé

        public int Dim => Assets.Count;

        public Basket(List<Asset> assets, double[] weights)
        {
            if (assets == null || assets.Count == 0) throw new ArgumentException("Assets required");
            if (weights == null || weights.Length != assets.Count) throw new ArgumentException("Weights size mismatch");

            Assets = assets;
            Weights = weights;

            // pas obligatoire, mais on vérifie un minimum
            double sum = 0.0;
            for (int i = 0; i < weights.Length; i++) sum += weights[i];
            if (Math.Abs(sum - 1.0) > 1e-6)
            {
                // on ne bloque pas, mais on prévient
                // (dans un vrai projet, tu pourrais normaliser)
            }
        }
    }

    public sealed class PriceResult
    {
        public double Price { get; set; }
        public double Variance { get; set; } // variance de l'estimateur (pas des payoffs)
    }
}
