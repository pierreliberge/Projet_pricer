using System;

namespace Projet_Pricer
{
    public static class MathUtils
    {
        // Box-Muller (gaussienne standard)
        public static double NextGaussian(Random rng)
        {
            // éviter log(0)
            double u1 = 1.0 - rng.NextDouble();
            double u2 = 1.0 - rng.NextDouble();
            return Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Cos(2.0 * Math.PI * u2);
        }

        // Décomposition de Cholesky (matrice symétrique définie positive)
        public static double[,] Cholesky(double[,] a)
        {
            int n = a.GetLength(0);
            if (n != a.GetLength(1)) throw new ArgumentException("Matrix must be square");

            double[,] l = new double[n, n];

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j <= i; j++)
                {
                    double sum = a[i, j];
                    for (int k = 0; k < j; k++)
                        sum -= l[i, k] * l[j, k];

                    if (i == j)
                    {
                        // petite régularisation si ça arrive proche de 0
                        if (sum <= 1e-14) sum = 1e-14;
                        l[i, j] = Math.Sqrt(sum);
                    }
                    else
                    {
                        l[i, j] = sum / l[j, j];
                    }
                }
            }

            return l;
        }

        public static double[] MulLowerTri(double[,] l, double[] v)
        {
            int n = l.GetLength(0);
            double[] res = new double[n];
            for (int i = 0; i < n; i++)
            {
                double s = 0.0;
                for (int j = 0; j <= i; j++)
                    s += l[i, j] * v[j];
                res[i] = s;
            }
            return res;
        }

        // CDF normale standard (approx)
        public static double NormCdf(double x)
        {
            // Abramowitz & Stegun-like approximation
            // suffisamment correct pour du pricing
            double t = 1.0 / (1.0 + 0.2316419 * Math.Abs(x));
            double d = 0.3989423 * Math.Exp(-0.5 * x * x);
            double prob = d * t * (0.3193815 + t * (-0.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274))));
            return x >= 0 ? 1.0 - prob : prob;
        }
    }
}