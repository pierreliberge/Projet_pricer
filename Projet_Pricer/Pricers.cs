using System;

namespace Projet_Pricer
{
    public static class BasketPricers
    {
        // ---------- Moment Matching (lognormale sur B_T via m1,m2) ----------
        public static double MomentMatchingPrice(IMarketModel model, OptionSpec opt)
        {
            int n = model.Dim;
            var b = model.Basket;
            double T = opt.Maturity;
            double df = model.DiscountFactor(T);

            // m1 = E[B_T] = sum w_i E[S_i(T)]
            double m1 = 0.0;
            for (int i = 0; i < n; i++)
            {
                double eSi = Math.Exp(model.LogMean(i, T) + 0.5 * model.LogVar(i, T)); // E[exp(lnSi)]
                m1 += b.Weights[i] * eSi;
            }

            // m2 = E[B_T^2] = sum_i sum_j w_i w_j E[S_i S_j]
            double m2 = 0.0;
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    double mean = model.LogMean(i, T) + model.LogMean(j, T);
                    double var = model.LogVar(i, T) + model.LogVar(j, T) + 2.0 * model.LogCov(i, j, T);
                    // E[Si*Sj] = E[exp(lnSi + lnSj)] = exp(mean + 0.5*var)
                    double eSij = Math.Exp(mean + 0.5 * var);
                    m2 += b.Weights[i] * b.Weights[j] * eSij;
                }
            }

            if (m1 <= 0) throw new Exception("m1 <= 0, check inputs");
            if (m2 <= m1 * m1) m2 = m1 * m1 * (1.0 + 1e-12);

            double sig2 = Math.Log(m2 / (m1 * m1));
            double sig = Math.Sqrt(sig2);

            double K = opt.Strike;

            // Prix call lognormal (sur B_T approx)
            double d1 = (Math.Log(m1 / K) + 0.5 * sig2) / sig;
            double d2 = d1 - sig;

            double call = df * (m1 * MathUtils.NormCdf(d1) - K * MathUtils.NormCdf(d2));
            if (opt.Type == OptionType.Call) return call;

            // Put via parité (sur l'approx lognormal)
            double put = call - df * (m1 - K);
            return put;
        }

        // ---------- Monte Carlo + control variate (géométrique analytique) ----------
        public static PriceResult MonteCarloPriceWithControlVariate(IMarketModel model, OptionSpec opt, int nPaths, int seed)
        {
            if (nPaths <= 1000) throw new ArgumentException("Use at least ~1000 paths");

            int n = model.Dim;
            var basket = model.Basket;
            double T = opt.Maturity;
            double K = opt.Strike;
            double df = model.DiscountFactor(T);

            // Cholesky sur la matrice des corrélation dans l'espace log
            // Attention: ici on reconstruit la matrice de corr "instantanée" via cov/vars => on utilise la cov finale
            // On fabrique une matrice CorrLog à partir des covariances log terminales.
            double[,] corrLog = new double[n, n];
            double[] std = new double[n];

            for (int i = 0; i < n; i++)
            {
                double v = model.LogVar(i, T);
                std[i] = Math.Sqrt(Math.Max(v, 1e-16));
            }
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    double cov = model.LogCov(i, j, T);
                    double denom = std[i] * std[j];
                    corrLog[i, j] = (denom < 1e-16) ? (i == j ? 1.0 : 0.0) : cov / denom;
                }
            }

            var L = MathUtils.Cholesky(corrLog);
            var rng = new Random(seed);

            // Control variate: option sur panier géométrique (analytique)
            double geoAnalytic = GeoBasketOptionPrice(model, opt);

            // Stats pour beta
            double sumY = 0.0, sumX = 0.0;
            double sumYY = 0.0, sumXX = 0.0, sumXY = 0.0;

            // On stocke aussi les samples pour variance finale (simple et clair)
            double[] adjSamples = new double[nPaths];

            // 1) on simule Y (arith payoff disc) et X (geo payoff disc)
            for (int p = 0; p < nPaths; p++)
            {
                double[] eps = new double[n];
                for (int i = 0; i < n; i++) eps[i] = MathUtils.NextGaussian(rng);

                double[] z = MathUtils.MulLowerTri(L, eps);

                // Simule ln S_i(T) = mean + std * z_i
                double basketArith = 0.0;

                // G = exp(sum w_i ln S_i) => on accumule directement le log
                double logG = 0.0;

                for (int i = 0; i < n; i++)
                {
                    double mean = model.LogMean(i, T);
                    double lnSi = mean + std[i] * z[i];
                    double Si = Math.Exp(lnSi);

                    basketArith += basket.Weights[i] * Si;
                    logG += basket.Weights[i] * lnSi;
                }

                double GT = Math.Exp(logG);

                double payoffArith = (opt.Type == OptionType.Call)
                    ? Math.Max(basketArith - K, 0.0)
                    : Math.Max(K - basketArith, 0.0);

                double payoffGeo = (opt.Type == OptionType.Call)
                    ? Math.Max(GT - K, 0.0)
                    : Math.Max(K - GT, 0.0);

                double Y = df * payoffArith;
                double X = df * payoffGeo;

                sumY += Y; sumX += X;
                sumYY += Y * Y;
                sumXX += X * X;
                sumXY += X * Y;

                // on remplit plus tard avec beta
                adjSamples[p] = Y; // temporaire
            }

            double meanY = sumY / nPaths;
            double meanX = sumX / nPaths;

            double varX = (sumXX / nPaths) - meanX * meanX;
            double covXY = (sumXY / nPaths) - meanX * meanY;

            double beta = 0.0;
            if (varX > 1e-18) beta = covXY / varX;

            // 2) on re-parcourt les samples "Y" stockés? ici on n'a pas stocké X.
            // donc on resimule proprement une 2ème fois pour construire l'estimateur ajusté et sa variance.
            // (oui c'est 2 passes, mais c'est clair, et ça marche. Tu pourras optimiser après.)

            rng = new Random(seed); // reset pour refaire les mêmes tirages

            double sumAdj = 0.0;
            double sumAdj2 = 0.0;

            for (int p = 0; p < nPaths; p++)
            {
                double[] eps = new double[n];
                for (int i = 0; i < n; i++) eps[i] = MathUtils.NextGaussian(rng);
                double[] z = MathUtils.MulLowerTri(L, eps);

                double basketArith = 0.0;
                double logG = 0.0;

                for (int i = 0; i < n; i++)
                {
                    double mean = model.LogMean(i, T);
                    double lnSi = mean + std[i] * z[i];
                    double Si = Math.Exp(lnSi);

                    basketArith += basket.Weights[i] * Si;
                    logG += basket.Weights[i] * lnSi;
                }

                double GT = Math.Exp(logG);

                double payoffArith = (opt.Type == OptionType.Call)
                    ? Math.Max(basketArith - K, 0.0)
                    : Math.Max(K - basketArith, 0.0);

                double payoffGeo = (opt.Type == OptionType.Call)
                    ? Math.Max(GT - K, 0.0)
                    : Math.Max(K - GT, 0.0);

                double Y = df * payoffArith;
                double X = df * payoffGeo;

                double adj = Y + beta * (geoAnalytic - X);

                sumAdj += adj;
                sumAdj2 += adj * adj;
            }

            double price = sumAdj / nPaths;

            // variance de l'estimateur = Var(sample) / N
            double varSample = (sumAdj2 / nPaths) - price * price;
            double varEstimator = varSample / nPaths;

            return new PriceResult { Price = price, Variance = varEstimator };
        }

        // ---------- Prix analytique de l'option sur panier géométrique ----------
        // ln G = sum w_i ln S_i => normal avec mean = sum w_i mean_i ; var = sum_i sum_j w_i w_j cov_ij
        private static double GeoBasketOptionPrice(IMarketModel model, OptionSpec opt)
        {
            int n = model.Dim;
            var b = model.Basket;
            double T = opt.Maturity;
            double K = opt.Strike;
            double df = model.DiscountFactor(T);

            double m = 0.0;
            for (int i = 0; i < n; i++)
                m += b.Weights[i] * model.LogMean(i, T);

            double v = 0.0;
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                    v += b.Weights[i] * b.Weights[j] * model.LogCov(i, j, T);

            v = Math.Max(v, 1e-16);
            double s = Math.Sqrt(v);

            // E[G_T] = exp(m + 0.5 v)
            double EG = Math.Exp(m + 0.5 * v);

            double d1 = (Math.Log(EG / K) + 0.5 * v) / s;
            double d2 = d1 - s;

            double call = df * (EG * MathUtils.NormCdf(d1) - K * MathUtils.NormCdf(d2));
            if (opt.Type == OptionType.Call) return call;

            double put = call - df * (EG - K);
            return put;
        }
    }
}
