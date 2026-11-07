using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Facturas
{
    internal sealed class Item
    {
        public string Desc = "";
        public string Qty = "1";
        public string Price = "";
    }

    internal sealed class Invoice
    {
        public string Id = "";
        public string Number = "";
        public string Date = "";
        public string Client = "";
        public string TaxId = "";
        public string Address = "";
        public List<Item> Items = new List<Item>();
        public string Discount = "0";
        public string Notes = "";
        public string Status = "Pendiente";
        public string PaidDate = "";
        public long Updated;
    }

    internal sealed class Db
    {
        public string Symbol = "L";
        public List<Invoice> Invoices = new List<Invoice>();
        public string Current;
        public List<string> Deleted = new List<string>();
    }

    internal struct Totals
    {
        public double Sub;
        public double Disc;
        public double Total;
    }

    internal static class U
    {
        public static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public static double Num(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            s = s.Trim();
            double d;
            if (!double.TryParse(s, NumberStyles.Float, Inv, out d))
            {
                if (!double.TryParse(s.Replace(',', '.'), NumberStyles.Float, Inv, out d)) return 0;
            }
            if (double.IsNaN(d) || double.IsInfinity(d)) return 0;
            return d;
        }

        public static double R2(double x)
        {
            return Math.Round(x, 2, MidpointRounding.AwayFromZero);
        }

        public static string Money(double n, string symbol)
        {
            string s = R2(n).ToString("N2", Inv);
            return string.IsNullOrEmpty(symbol) ? s : symbol + " " + s;
        }

        public static string Uid()
        {
            return Guid.NewGuid().ToString("N").Substring(0, 12);
        }

        public static long NowMs()
        {
            return DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        }

        public static string Today()
        {
            return DateTime.Now.ToString("yyyy-MM-dd", Inv);
        }

        public static string FmtDate(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var p = s.Split('-');
            return p.Length == 3 ? p[2] + "/" + p[1] + "/" + p[0] : s;
        }

        public static DateTime ParseDate(string s)
        {
            DateTime d;
            if (DateTime.TryParseExact(s ?? "", "yyyy-MM-dd", Inv, DateTimeStyles.None, out d)) return d;
            return DateTime.Today;
        }

        public static Item BlankItem()
        {
            return new Item();
        }

        public static Totals Calc(Invoice inv)
        {
            double sub = inv.Items.Sum(it => Num(it.Qty) * Num(it.Price));
            double disc = sub * Num(inv.Discount) / 100.0;
            return new Totals { Sub = sub, Disc = disc, Total = sub - disc };
        }

        public static bool IsBlank(Invoice inv)
        {
            return string.IsNullOrEmpty(inv.Client) && string.IsNullOrEmpty(inv.TaxId) &&
                   string.IsNullOrEmpty(inv.Address) &&
                   inv.Items.All(it => string.IsNullOrEmpty(it.Desc) && !(Num(it.Price) > 0));
        }

        public static string NextNumber(Db db)
        {
            int max = 0;
            foreach (var i in db.Invoices)
            {
                var m = Regex.Match(i.Number ?? "", @"(\d+)\s*$");
                if (m.Success)
                {
                    int n;
                    if (int.TryParse(m.Groups[1].Value, out n)) max = Math.Max(max, n);
                }
            }
            return (max + 1).ToString().PadLeft(3, '0');
        }

        public static Invoice NewInvoice(Db db)
        {
            var inv = new Invoice
            {
                Id = Uid(),
                Number = NextNumber(db),
                Date = Today(),
                Updated = NowMs()
            };
            inv.Items.Add(BlankItem());
            inv.Items.Add(BlankItem());
            inv.Items.Add(BlankItem());
            return inv;
        }

        private static int NumberOf(string s)
        {
            int n;
            return int.TryParse(Regex.Replace(s ?? "", @"\D", ""), out n) ? n : 0;
        }

        public static List<Invoice> Sorted(IEnumerable<Invoice> list)
        {
            return list
                .OrderByDescending(i => i.Date ?? "", StringComparer.Ordinal)
                .ThenByDescending(i => NumberOf(i.Number))
                .ToList();
        }

        public static void Merge(Db db, Db other)
        {
            var map = new Dictionary<string, Invoice>();
            foreach (var i in other.Invoices) map[i.Id] = i;
            var gone = new HashSet<string>(other.Deleted.Concat(db.Deleted));
            foreach (var i in db.Invoices)
            {
                Invoice o;
                if (!map.TryGetValue(i.Id, out o) || i.Updated > o.Updated) map[i.Id] = i;
            }
            db.Invoices = map.Values.Where(i => !gone.Contains(i.Id)).ToList();
            db.Deleted = gone.ToList();
        }
    }
}
