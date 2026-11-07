using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace Facturas
{
    /// <summary>
    /// Guarda las facturas en facturas.json (mismo formato que la versión web, así que
    /// los respaldos de una sirven en la otra) y una copia diaria en la subcarpeta "respaldo".
    /// </summary>
    internal static class Storage
    {
        private const string DataFile = "facturas.json";
        private const string BackupDir = "respaldo";

        private static JavaScriptSerializer Ser()
        {
            return new JavaScriptSerializer { MaxJsonLength = int.MaxValue, RecursionLimit = 200 };
        }

        // ---------- carpeta elegida ----------
        private static string ConfigPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Facturas", "config.json");
        }

        public static string DefaultFolder()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Facturas");
        }

        public static string FallbackFolder()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Facturas", "datos");
        }

        public static string ReadFolder()
        {
            try
            {
                string cfg = ConfigPath();
                if (File.Exists(cfg))
                {
                    var d = Ser().DeserializeObject(File.ReadAllText(cfg, Encoding.UTF8)) as Dictionary<string, object>;
                    object f;
                    if (d != null && d.TryGetValue("folder", out f))
                    {
                        string path = Convert.ToString(f, U.Inv);
                        if (!string.IsNullOrWhiteSpace(path)) return path;
                    }
                }
            }
            catch { }
            return DefaultFolder();
        }

        public static void SaveFolder(string folder)
        {
            try
            {
                string cfg = ConfigPath();
                Directory.CreateDirectory(Path.GetDirectoryName(cfg));
                File.WriteAllText(cfg, Ser().Serialize(new Dictionary<string, object> { { "folder", folder } }),
                    new UTF8Encoding(false));
            }
            catch { }
        }

        // ---------- JSON ----------
        private static string Str(object o)
        {
            return o == null ? "" : Convert.ToString(o, U.Inv);
        }

        private static IEnumerable Seq(object o)
        {
            return o as IEnumerable;
        }

        public static string ToJson(Db db)
        {
            var invoices = db.Invoices.Select(i => new Dictionary<string, object>
            {
                { "id", i.Id },
                { "number", i.Number },
                { "date", i.Date },
                { "client", i.Client },
                { "taxId", i.TaxId },
                { "address", i.Address },
                { "items", i.Items.Select(it => new Dictionary<string, object>
                    {
                        { "desc", it.Desc }, { "qty", it.Qty }, { "price", it.Price }
                    }).ToList() },
                { "discount", i.Discount },
                { "notes", i.Notes },
                { "status", i.Status },
                { "paidDate", i.PaidDate },
                { "updated", i.Updated }
            }).ToList();

            var root = new Dictionary<string, object>
            {
                { "settings", new Dictionary<string, object> { { "symbol", db.Symbol } } },
                { "invoices", invoices },
                { "current", db.Current },
                { "deleted", db.Deleted }
            };
            return Ser().Serialize(root);
        }

        public static Db FromJson(string text)
        {
            var root = Ser().DeserializeObject(text) as Dictionary<string, object>;
            if (root == null || !root.ContainsKey("invoices")) throw new FormatException("formato");
            var arr = Seq(root["invoices"]);
            if (arr == null) throw new FormatException("formato");

            var db = new Db();
            foreach (object o in arr)
            {
                var d = o as Dictionary<string, object>;
                if (d == null) continue;
                var inv = new Invoice
                {
                    Id = Str(Get(d, "id")),
                    Number = Str(Get(d, "number")),
                    Date = Str(Get(d, "date")),
                    Client = Str(Get(d, "client")),
                    TaxId = Str(Get(d, "taxId")),
                    Address = Str(Get(d, "address")),
                    Discount = Str(Get(d, "discount")),
                    Notes = Str(Get(d, "notes")),
                    Status = Str(Get(d, "status")),
                    PaidDate = Str(Get(d, "paidDate"))
                };
                if (string.IsNullOrEmpty(inv.Id)) inv.Id = U.Uid();
                if (inv.Status != "Pagada") inv.Status = "Pendiente";
                if (string.IsNullOrEmpty(inv.Discount)) inv.Discount = "0";
                try { inv.Updated = Convert.ToInt64(Get(d, "updated") ?? 0, U.Inv); } catch { inv.Updated = 0; }

                var its = Seq(Get(d, "items"));
                if (its != null)
                {
                    foreach (object io in its)
                    {
                        var id = io as Dictionary<string, object>;
                        if (id == null) continue;
                        inv.Items.Add(new Item
                        {
                            Desc = Str(Get(id, "desc")),
                            Qty = Str(Get(id, "qty")),
                            Price = Str(Get(id, "price"))
                        });
                    }
                }
                if (inv.Items.Count == 0) inv.Items.Add(U.BlankItem());
                db.Invoices.Add(inv);
            }

            var settings = Get(root, "settings") as Dictionary<string, object>;
            if (settings != null && settings.ContainsKey("symbol")) db.Symbol = Str(settings["symbol"]);
            db.Current = Str(Get(root, "current"));
            var del = Seq(Get(root, "deleted"));
            if (del != null) foreach (object x in del) db.Deleted.Add(Str(x));
            return db;
        }

        private static object Get(Dictionary<string, object> d, string key)
        {
            object v;
            return d.TryGetValue(key, out v) ? v : null;
        }

        // ---------- disco ----------
        /// <summary>Devuelve null si no hay archivo (o si estaba dañado: en ese caso notice explica qué pasó).</summary>
        public static Db Load(string folder, out string notice)
        {
            notice = null;
            string file = Path.Combine(folder, DataFile);
            if (!File.Exists(file)) return null;
            try
            {
                return FromJson(File.ReadAllText(file, Encoding.UTF8));
            }
            catch
            {
                string bad = "facturas_danado_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", U.Inv) + ".json";
                try { File.Copy(file, Path.Combine(folder, bad), true); } catch { }

                string bk = Path.Combine(folder, BackupDir);
                if (Directory.Exists(bk))
                {
                    foreach (string f in Directory.GetFiles(bk, "facturas_*.json")
                        .OrderByDescending(x => x, StringComparer.OrdinalIgnoreCase))
                    {
                        try
                        {
                            var d = FromJson(File.ReadAllText(f, Encoding.UTF8));
                            notice = "El archivo de facturas estaba dañado. Guardé una copia como «" + bad +
                                     "» y recuperé tus datos desde el respaldo «" + Path.GetFileName(f) + "».";
                            return d;
                        }
                        catch { }
                    }
                }
                notice = "El archivo de facturas estaba dañado y no encontré un respaldo válido. " +
                         "Guardé una copia como «" + bad + "» en la carpeta de facturas.";
                return null;
            }
        }

        private static void WriteAtomic(string path, string text)
        {
            string tmp = path + ".tmp";
            File.WriteAllText(tmp, text, new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(tmp, path, null);
            else File.Move(tmp, path);
        }

        public static void Save(string folder, Db db)
        {
            Directory.CreateDirectory(folder);
            string text = ToJson(db);
            WriteAtomic(Path.Combine(folder, DataFile), text);
            string bk = Path.Combine(folder, BackupDir);
            Directory.CreateDirectory(bk);
            WriteAtomic(Path.Combine(bk, "facturas_" + U.Today() + ".json"), text);
        }
    }
}
