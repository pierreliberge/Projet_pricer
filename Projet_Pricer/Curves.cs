using System;
using System.Collections.Generic;

namespace Projet_Pricer
{
    // Courbe de taux: on stocke des discount factors DF(t) et on interpole log(DF)
    public sealed class RateCurve
    {
        private readonly List<(double t, double df)> _points;

        private RateCurve(List<(double t, double df)> points)
        {
            _points = points;
            _points.Sort((a, b) => a.t.CompareTo(b.t));
        }

        public static RateCurve FromZeroRates(List<(double t, double z)> zeroRates)
        {
            var pts = new List<(double t, double df)>();
            foreach (var p in zeroRates)
            {
                if (p.t <= 0) throw new ArgumentException("Maturities must be > 0");
                double df = Math.Exp(-p.z * p.t);
                pts.Add((p.t, df));
            }
            return new RateCurve(pts);
        }

        public double DF(double t)
        {
            if (t <= 0) return 1.0;

            if (t <= _points[0].t) return InterpLogDF(t, 0, 1);
            if (t >= _points[_points.Count - 1].t) return InterpLogDF(t, _points.Count - 2, _points.Count - 1);

            for (int i = 0; i < _points.Count - 1; i++)
            {
                if (t >= _points[i].t && t <= _points[i + 1].t)
                    return InterpLogDF(t, i, i + 1);
            }

            // fallback
            return _points[_points.Count - 1].df;
        }

        public double IntegralR(double t)
        {
            // -ln DF = ∫ r(u) du (dans ce modèle basé DF)
            double df = DF(t);
            return -Math.Log(df);
        }

        private double InterpLogDF(double t, int i0, int i1)
        {
            var p0 = _points[i0];
            var p1 = _points[i1];

            if (Math.Abs(p1.t - p0.t) < 1e-12) return p0.df;

            double w = (t - p0.t) / (p1.t - p0.t);
            double logdf = (1.0 - w) * Math.Log(p0.df) + w * Math.Log(p1.df);
            return Math.Exp(logdf);
        }
    }

    // Vol(t): interpolation linéaire sur (t, vol), intégrales trapezoidales
    public sealed class VolCurve
    {
        private readonly List<(double t, double vol)> _points;

        public VolCurve(List<(double t, double vol)> points)
        {
            _points = points ?? throw new ArgumentNullException(nameof(points));
            _points.Sort((a, b) => a.t.CompareTo(b.t));
            if (_points.Count < 2) throw new ArgumentException("VolCurve needs at least 2 points");
        }

        public double Vol(double t)
        {
            if (t <= _points[0].t) return _points[0].vol;
            if (t >= _points[_points.Count - 1].t) return _points[_points.Count - 1].vol;

            for (int i = 0; i < _points.Count - 1; i++)
            {
                if (t >= _points[i].t && t <= _points[i + 1].t)
                {
                    var p0 = _points[i];
                    var p1 = _points[i + 1];
                    double w = (t - p0.t) / (p1.t - p0.t);
                    return (1.0 - w) * p0.vol + w * p1.vol;
                }
            }
            return _points[_points.Count - 1].vol;
        }

        public double IntegralVol2(double T, int steps = 200)
        {
            if (T <= 0) return 0.0;

            double dt = T / steps;
            double sum = 0.0;
            double t0 = 0.0;
            double v0 = Vol(t0);

            for (int k = 1; k <= steps; k++)
            {
                double t1 = k * dt;
                double v1 = Vol(t1);
                sum += 0.5 * (v0 * v0 + v1 * v1) * (t1 - t0);
                t0 = t1;
                v0 = v1;
            }
            return sum;
        }

        public static double IntegralVolProduct(VolCurve a, VolCurve b, double T, int steps = 200)
        {
            if (T <= 0) return 0.0;

            double dt = T / steps;
            double sum = 0.0;
            double t0 = 0.0;
            double va0 = a.Vol(t0);
            double vb0 = b.Vol(t0);

            for (int k = 1; k <= steps; k++)
            {
                double t1 = k * dt;
                double va1 = a.Vol(t1);
                double vb1 = b.Vol(t1);

                sum += 0.5 * (va0 * vb0 + va1 * vb1) * (t1 - t0);

                t0 = t1;
                va0 = va1;
                vb0 = vb1;
            }
            return sum;
        }
    }
}
