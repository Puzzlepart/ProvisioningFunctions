using System;
using System.Drawing;

namespace Pzl.ProvisioningFunctions.Helpers
{
    internal class HSLColor
    {
        public const int MaxRgb = 255;
        public static readonly HSLColor Black = new HSLColor(0.6666667f, 0.0f, 0.0f);
        public static readonly HSLColor White = new HSLColor(0.6666667f, 0.0f, 1f);
        private readonly byte _alpha;
        private float _hue;
        private float _luminance;
        private float _saturation;

        public HSLColor(float hue, float saturation, float luminance)
            : this(hue, saturation, luminance, byte.MaxValue)
        {
        }

        public HSLColor(float hue, float saturation, float luminance, byte alpha)
        {
            _hue = 0.0f;
            _saturation = 0.0f;
            _luminance = 0.0f;
            _alpha = alpha;
            Hue = hue;
            Saturation = saturation;
            Luminance = luminance;
        }

        public float Hue
        {
            get => _hue;
            set => _hue = value == -1.0 ? -1f : Math.Min(1f, Math.Max(0.0f, value));
        }

        public float Saturation
        {
            get => _saturation;
            set => _saturation = Math.Min(1f, Math.Max(0.0f, value));
        }

        public float Luminance
        {
            get => _luminance;
            set => _luminance = Math.Min(1f, Math.Max(0.0f, value));
        }

        public Color ToRgbColor()
        {
            float num1;
            float num2;
            float num3;
            if (FloatEquals(_saturation, 0.0f))
            {
                num1 = _luminance;
                num2 = _luminance;
                num3 = _luminance;
            }
            else
            {
                float m2 = (double)_luminance > 0.5
                    ? (float)(_luminance + (double)_saturation - _luminance * (double)_saturation)
                    : _luminance * (1f + _saturation);
                float m1 = 2f * _luminance - m2;
                num1 = HueToRgb(m1, m2, _hue + 0.3333333f);
                num2 = HueToRgb(m1, m2, _hue);
                num3 = HueToRgb(m1, m2, _hue - 0.3333333f);
            }
            return Color.FromArgb(_alpha, (int)(num1 * (double)byte.MaxValue), (int)(num2 * (double)byte.MaxValue),
                (int)(num3 * (double)byte.MaxValue));
        }

        private static float HueToRgb(float m1, float m2, float hue)
        {
            if (hue < 0.0)
                ++hue;
            if (hue > 1.0)
                --hue;
            if (6.0 * hue < 1.0)
                return m1 + (float)((m2 - (double)m1) * hue * 6.0);
            if (2.0 * hue < 1.0)
                return m2;
            if (3.0 * hue < 2.0)
                return m1 + (float)((m2 - (double)m1) * (0.666666686534882 - hue) * 6.0);
            return m1;
        }

        public static HSLColor FromRgbColor(Color color)
        {
            float num1 = color.R / (float)byte.MaxValue;
            float num2 = color.G / (float)byte.MaxValue;
            float val2 = color.B / (float)byte.MaxValue;
            float num3 = Math.Max(Math.Max(num1, num2), val2);
            float n2 = Math.Min(Math.Min(num1, num2), val2);
            var luminance = (float)((num3 + (double)n2) / 2.0);
            float saturation;
            float hue;
            if (FloatEquals(num3, n2))
            {
                saturation = 0.0f;
                hue = -1f;
            }
            else
            {
                float num4 = num3 - n2;
                float num5 = num3 + n2;
                saturation = (double)luminance > 0.5 ? num4 / (2f - num5) : num4 / num5;
                float num6 = (float)((num3 - (double)num1) * 0.16666667163372) / num4;
                float num7 = (float)((num3 - (double)num2) * 0.16666667163372) / num4;
                float num8 = (float)((num3 - (double)val2) * 0.16666667163372) / num4;
                hue = !FloatEquals(num1, num3)
                    ? (!FloatEquals(num2, num3) ? 0.6666667f + num7 - num6 : 0.3333333f + num6 - num8)
                    : num8 - num7;
                if (hue < 0.0)
                    ++hue;
                if (hue > 1.0)
                    --hue;
            }
            return new HSLColor(hue, saturation, luminance, color.A);
        }

        public static bool FloatEquals(float n1, float n2)
        {
            return Math.Abs(n1 - n2) < 2.80259692864963E-45;
        }

        public static int GetLuminance(Color color)
        {
            return (color.R * 13927 + color.G * 46885 + color.B * 4725) / 65536;
        }

        public void Lighten(float factor)
        {
            Luminance = (float)(Luminance * (double)factor + (1.0 - factor));
        }

        public void Darken(float factor)
        {
            Luminance = Luminance * factor;
        }
    }
}
