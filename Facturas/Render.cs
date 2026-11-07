using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Linq;

namespace Facturas
{
    /// <summary>Dibuja la factura (para imprimir, guardar como PDF o copiar como imagen).</summary>
    internal static class Render
    {
        public const float W = 900f;
        private const float Pad = 48f;

        private static readonly Color Navy = Color.FromArgb(0x1f, 0x38, 0x64);
        private static readonly Color Grey = Color.FromArgb(0x6b, 0x72, 0x80);
        private static readonly Color Line = Color.FromArgb(0xdf, 0xe3, 0xea);
        private static readonly Color Ink = Color.FromArgb(0x1b, 0x1f, 0x27);

        private static readonly Dictionary<string, Font> Fonts = new Dictionary<string, Font>();

        private static Font F(float px, bool bold)
        {
            string key = px.ToString(U.Inv) + (bold ? "b" : "r");
            Font f;
            if (!Fonts.TryGetValue(key, out f))
            {
                f = new Font("Segoe UI", px, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
                Fonts[key] = f;
            }
            return f;
        }

        /// <summary>
        /// Dibuja la factura en g (si g es null solo mide) y devuelve el alto total.
        /// mg es un Graphics en píxeles que se usa siempre para medir el texto.
        /// </summary>
        public static float Draw(Graphics g, Graphics mg, Invoice inv, string sym)
        {
            bool paint = g != null;

            var sf = (StringFormat)StringFormat.GenericTypographic.Clone();
            sf.FormatFlags |= StringFormatFlags.NoWrap;
            var sfR = (StringFormat)sf.Clone();
            sfR.Alignment = StringAlignment.Far;

            void Box(Color c, float x, float y, float w, float h)
            {
                if (!paint) return;
                using (var b = new SolidBrush(c)) g.FillRectangle(b, x, y, w, h);
            }

            void L(string s, float x, float y, Font f, Color c)
            {
                if (!paint || string.IsNullOrEmpty(s)) return;
                using (var b = new SolidBrush(c)) g.DrawString(s, f, b, x, y, sf);
            }

            void R(string s, float xRight, float y, Font f, Color c)
            {
                if (!paint || string.IsNullOrEmpty(s)) return;
                using (var b = new SolidBrush(c))
                    g.DrawString(s, f, b, new RectangleF(xRight - 600f, y, 600f, f.Height + 6f), sfR);
            }

            List<string> Wrap(string text, float maxW, Font f)
            {
                var res = new List<string>();
                foreach (string par in (text ?? "").Replace("\r", "").Split('\n'))
                {
                    string line = "";
                    foreach (string word in par.Split(' '))
                    {
                        string t = line.Length > 0 ? line + " " + word : word;
                        if (line.Length > 0 && mg.MeasureString(t, f, PointF.Empty, sf).Width > maxW)
                        {
                            res.Add(line);
                            line = word;
                        }
                        else line = t;
                    }
                    res.Add(line);
                }
                return res;
            }

            var items = inv.Items.Where(it => !string.IsNullOrEmpty(it.Desc) || U.Num(it.Price) > 0).ToList();
            Totals c2 = U.Calc(inv);

            float y0 = Pad;

            // ---- título ----
            L("FACTURA", Pad, y0, F(36, true), Navy);
            R("No. " + inv.Number, W - Pad, y0 + 12, F(20, true), Grey);
            y0 += 56;
            Box(Navy, Pad, y0, W - 2 * Pad, 3);
            y0 += 22;

            // ---- campos ----
            float colW = (W - 2 * Pad - 48f) / 3f;
            void FieldRow(string[][] fields)
            {
                float h = 0;
                for (int i = 0; i < fields.Length; i++)
                {
                    float x = Pad + i * (colW + 24f);
                    L(fields[i][0].ToUpperInvariant(), x, y0, F(11, true), Grey);
                    var lines = Wrap(fields[i][1], colW, F(16, false));
                    for (int k = 0; k < lines.Count; k++) L(lines[k], x, y0 + 18 + k * 21, F(16, false), Ink);
                    h = Math.Max(h, 18 + lines.Count * 21);
                }
                y0 += h + 10;
            }

            string estado = inv.Status + (inv.Status == "Pagada" && !string.IsNullOrEmpty(inv.PaidDate)
                ? " · " + U.FmtDate(inv.PaidDate) : "");
            FieldRow(new[]
            {
                new[] { "Factura No.", inv.Number },
                new[] { "Fecha", U.FmtDate(inv.Date) },
                new[] { "Estado", estado }
            });
            FieldRow(new[]
            {
                new[] { "Cliente", inv.Client },
                new[] { "RTN / NIT", inv.TaxId },
                new[] { "Dirección", inv.Address }
            });
            y0 += 8;

            // ---- tabla ----
            float xN = Pad + 12, xD = Pad + 44, rQ = W - Pad - 290, rP = W - Pad - 160, rS = W - Pad - 12;
            float descW = rQ - 80 - xD;
            Box(Navy, Pad, y0, W - 2 * Pad, 34);
            L("#", xN, y0 + 9, F(12, true), Color.White);
            L("DESCRIPCIÓN", xD, y0 + 9, F(12, true), Color.White);
            R("CANT.", rQ, y0 + 9, F(12, true), Color.White);
            R("PRECIO", rP, y0 + 9, F(12, true), Color.White);
            R("SUBTOTAL", rS, y0 + 9, F(12, true), Color.White);
            y0 += 34;

            for (int i = 0; i < items.Count; i++)
            {
                var it = items[i];
                var lines = Wrap(it.Desc, descW, F(15, false));
                float h = Math.Max(38f, lines.Count * 21f + 16f);
                L((i + 1).ToString(), xN, y0 + 8, F(15, false), Grey);
                for (int k = 0; k < lines.Count; k++) L(lines[k], xD, y0 + 8 + k * 21, F(15, false), Ink);
                R(U.Num(it.Qty).ToString(U.Inv), rQ, y0 + 8, F(15, false), Ink);
                R(string.IsNullOrEmpty(it.Price) ? "" : U.Money(U.Num(it.Price), sym), rP, y0 + 8, F(15, false), Ink);
                R(U.Money(U.Num(it.Qty) * U.Num(it.Price), sym), rS, y0 + 8, F(15, true), Ink);
                y0 += h;
                Box(Line, Pad, y0 - 1, W - 2 * Pad, 1);
            }
            y0 += 24;

            // ---- totales ----
            void TRow(string label, string value)
            {
                R(label, W - Pad - 160, y0, F(15, true), Grey);
                R(value, W - Pad - 12, y0, F(15, true), Ink);
                y0 += 28;
            }
            TRow("Subtotal", U.Money(c2.Sub, sym));
            if (U.Num(inv.Discount) > 0)
                TRow("Descuento (" + U.Num(inv.Discount).ToString(U.Inv) + "%)", "- " + U.Money(c2.Disc, sym));
            y0 += 6;
            Box(Navy, W - Pad - 330, y0, 330, 46);
            L("TOTAL", W - Pad - 316, y0 + 12, F(18, true), Color.White);
            R(U.Money(c2.Total, sym), W - Pad - 12, y0 + 10, F(21, true), Color.White);
            y0 += 46;

            // ---- notas ----
            if (!string.IsNullOrWhiteSpace(inv.Notes))
            {
                y0 += 30;
                L("NOTAS / CONDICIONES DE PAGO", Pad, y0, F(11, true), Grey);
                var nl = Wrap(inv.Notes, W - 2 * Pad, F(15, false));
                for (int k = 0; k < nl.Count; k++) L(nl[k], Pad, y0 + 22 + k * 21, F(15, false), Ink);
                y0 += 22 + nl.Count * 21;
            }

            return y0 + Pad;
        }

        /// <summary>Crea la imagen de la factura. scale = 2 da buena calidad para WhatsApp; 3 para imprimir.</summary>
        public static Bitmap ToBitmap(Invoice inv, string sym, int scale)
        {
            using (var scratch = new Bitmap(1, 1))
            using (var mg = Graphics.FromImage(scratch))
            {
                mg.PageUnit = GraphicsUnit.Pixel;
                float h = Draw(null, mg, inv, sym);
                var bmp = new Bitmap((int)(W * scale), (int)Math.Ceiling(h * scale), PixelFormat.Format24bppRgb);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                    g.ScaleTransform(scale, scale);
                    Draw(g, mg, inv, sym);
                }
                return bmp;
            }
        }
    }
}
