using System;
using System.Collections.Generic;

namespace Projet_Pricer
{
    public interface IMarketModel
    {
        Basket Basket { get; }
        int Dim { get; }

        double DiscountFactor(double T);

        // moments/log-diffusion "équivalents" à l'échéance
        double LogMean(int i, double T);                  // E[ln S_i(T)]
        double LogVar(int i, double T);                   // Var[ln S_i(T)]
        double LogCov(int i, int j, double T);            // Cov[ln S_i(T), ln S_j(T)]
    }

    // H1: r const, q const, sigma const, corr const
    public sealed class MarketModelH1 : IMarketModel
    {
        public Basket Basket { get; }
        public int Dim => Basket.Dim;

        private readonly double _r;
        private readonly double[,] _corr;

        public MarketModelH1(Basket basket, double r, double[,] corr)
        {
            Basket = basket ?? throw new ArgumentNullException(nameof(basket));
            _r = r;
            _corr = corr ?? throw new ArgumentNullException(nameof(corr));
            if (_corr.GetLength(0) != Dim || _corr.GetLength(1) != Dim)
                throw new ArgumentException("Corr dimension mismatch");
        }

        public double DiscountFactor(double T) => Math.Exp(-_r * T);

        public double LogMean(int i, double T)
        {
            var a = Basket.Assets[i];
            double mu = (_r - a.DividendYield - 0.5 * a.VolConst * a.VolConst) * T;
            return Math.Log(a.Spot) + mu;
        }

        public double LogVar(int i, double T)
        {
            var a = Basket.Assets[i];
            return a.VolConst * a.VolConst * T;
        }

        public double LogCov(int i, int j, double T)
        {
            var ai = Basket.Assets[i];
            var aj = Basket.Assets[j];
            return _corr[i, j] * ai.VolConst * aj.VolConst * T;
        }
    }

    // H2: r(t) via DF curve, q const, sigma(t) via VolCurve, corr const
    public sealed class MarketModelH2 : IMarketModel
    {
        public Basket Basket { get; }
        public int Dim => Basket.Dim;

        private readonly RateCurve _rateCurve;
        private readonly Dictionary<string, VolCurve> _volCurves;
        private readonly double[,] _corr;

        public MarketModelH2(Basket basket, RateCurve rateCurve, Dictionary<string, VolCurve> volCurves, double[,] corr)
        {
            Basket = basket ?? throw new ArgumentNullException(nameof(basket));
            _rateCurve = rateCurve ?? throw new ArgumentNullException(nameof(rateCurve));
            _volCurves = volCurves ?? throw new ArgumentNullException(nameof(volCurves));
            _corr = corr ?? throw new ArgumentNullException(nameof(corr));

            if (_corr.GetLength(0) != Dim || _corr.GetLength(1) != Dim)
                throw new ArgumentException("Corr dimension mismatch");

            // check vols exist
            for (int i = 0; i < Dim; i++)
            {
                var tkr = Basket.Assets[i].Ticker;
                if (!_volCurves.ContainsKey(tkr))
                    throw new ArgumentException($"Missing vol curve for {tkr}");
            }
        }

        public double DiscountFactor(double T) => _rateCurve.DF(T);

        public double LogMean(int i, double T)
        {
            var a = Basket.Assets[i];
            var vol = _volCurves[a.Ticker];

            double intR = _rateCurve.IntegralR(T);
            double intSig2 = vol.IntegralVol2(T);

            // q constant => ∫ q dt = qT
            double mu = (intR - a.DividendYield * T) - 0.5 * intSig2;
            return Math.Log(a.Spot) + mu;
        }

        public double LogVar(int i, double T)
        {
            var a = Basket.Assets[i];
            var vol = _volCurves[a.Ticker];
            return vol.IntegralVol2(T);
        }

        public double LogCov(int i, int j, double T)
        {
            if (i == j) return LogVar(i, T);

            var ai = Basket.Assets[i];
            var aj = Basket.Assets[j];

            var voli = _volCurves[ai.Ticker];
            var volj = _volCurves[aj.Ticker];

            double intProd = VolCurve.IntegralVolProduct(voli, volj, T);
            return _corr[i, j] * intProd;
        }
    }
}
