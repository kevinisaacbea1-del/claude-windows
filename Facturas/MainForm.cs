using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Facturas
{
    internal sealed class MainForm : Form
    {
        // ---------- colores y escala ----------
        private static readonly Color Navy = Color.FromArgb(0x1f, 0x38, 0x64);
        private static readonly Color BgC = Color.FromArgb(0xf4, 0xf5, 0xf8);
        private static readonly Color LineC = Color.FromArgb(0xdf, 0xe3, 0xea);
        private static readonly Color Muted = Color.FromArgb(0x6b, 0x72, 0x80);
        private static readonly Color Ink = Color.FromArgb(0x1b, 0x1f, 0x27);
        private static readonly Color OkC = Color.FromArgb(0x2e, 0x7d, 0x32);
        private static readonly Color WarnC = Color.FromArgb(0xb4, 0x53, 0x09);
        private static readonly Color Cell = Color.FromArgb(0xff, 0xfb, 0xe6);
        private static readonly float K;

        static MainForm()
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) K = g.DpiX / 96f;
        }

        private static int S(int v) { return (int)Math.Round(v * K); }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        // ---------- estado ----------
        private Db db;
        private Invoice cur;
        private string folder;
        private string startNotice;
        private bool loading, regLoading, errShown;
        private readonly System.Windows.Forms.Timer saveTimer = new System.Windows.Forms.Timer { Interval = 400 };

        private static readonly string[] Symbols = { "L", "Q", "$", "" };
        private readonly Font statusFont = new Font("Segoe UI", 9.5f, FontStyle.Bold);

        // ---------- controles ----------
        private ComboBox cbSymbol, cbStatus, cbFilter;
        private Button btnCsv, btnBackup, btnImport, btnFolder, btnOpenFolder;
        private Button btnNew, btnDup, btnPrint, btnCopyImg, btnSaveImg, btnDel, btnAddLine, btnDelLine;
        private Label lblFolder, lblSaved, lblHdrNo, lblSub, lblDisc, lblTotal;
        private Label sTotal, sPaid, sPend;
        private TextBox tbNumber, tbClient, tbTax, tbAddress, tbNotes, tbDiscount, tbSearch;
        private DateTimePicker dtDate, dtPaid;
        private Panel fldPaid;
        private ListView lv;
        private DataGridView grid;

        public MainForm()
        {
            Text = "Facturas";
            Font = new Font("Segoe UI", 9.5f);
            BackColor = BgC;
            StartPosition = FormStartPosition.CenterScreen;
            var wa = Screen.PrimaryScreen.WorkingArea;
            Size = new Size(Math.Min(S(1320), (int)(wa.Width * 0.96)), Math.Min(S(820), (int)(wa.Height * 0.96)));
            MinimumSize = new Size(Math.Min(S(1100), wa.Width), Math.Min(S(700), wa.Height));

            BuildUi();
            StartData();
            HookEvents();
            RefreshAll();

            saveTimer.Tick += (s, e) => { saveTimer.Stop(); SaveNow(false); };
            Shown += (s, e) =>
            {
                if (startNotice != null)
                    MessageBox.Show(this, startNotice, "Facturas", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            FormClosing += (s, e) => { SaveNow(true); };
        }

        // =====================================================================
        //  Interfaz
        // =====================================================================
        private Button MakeBtn(string text, bool primary, bool danger)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlatStyle = FlatStyle.Flat,
                Padding = new Padding(S(6), S(2), S(6), S(2)),
                Margin = new Padding(0, 0, S(8), 0),
                MinimumSize = new Size(0, S(32)),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
                BackColor = primary ? Navy : Color.White,
                ForeColor = primary ? Color.White : (danger ? Color.FromArgb(0xb4, 0x23, 0x18) : Ink)
            };
            b.FlatAppearance.BorderColor = primary ? Navy : LineC;
            return b;
        }

        private Button GhostBtn(string text)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlatStyle = FlatStyle.Flat,
                Padding = new Padding(S(6), S(2), S(6), S(2)),
                Margin = new Padding(0, 0, S(8), 0),
                MinimumSize = new Size(0, S(32)),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
                BackColor = Navy,
                ForeColor = Color.White
            };
            b.FlatAppearance.BorderColor = Color.FromArgb(0x5a, 0x74, 0xa8);
            return b;
        }

        private Panel Field(string label, Control c)
        {
            var p = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, S(12), 0) };
            c.Dock = DockStyle.Top;
            var l = new Label
            {
                Text = label.ToUpperInvariant(),
                Dock = DockStyle.Top,
                Height = S(20),
                ForeColor = Muted,
                Font = new Font("Segoe UI", 8f, FontStyle.Bold)
            };
            p.Controls.Add(c);
            p.Controls.Add(l);
            return p;
        }

        private TableLayoutPanel Row(int height, params Control[] cells)
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = S(height),
                ColumnCount = cells.Length,
                RowCount = 1,
                Margin = Padding.Empty
            };
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            for (int i = 0; i < cells.Length; i++)
            {
                t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / cells.Length));
                t.Controls.Add(cells[i], i, 0);
            }
            return t;
        }

        private void BuildUi()
        {
            SuspendLayout();
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = Padding.Empty
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, S(56)));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, S(42)));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(BuildTop(), 0, 0);
            root.Controls.Add(BuildFolderBar(), 0, 1);
            root.Controls.Add(BuildContent(), 0, 2);
            Controls.Add(root);
            ResumeLayout(true);
        }

        private Control BuildTop()
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = Navy, Margin = Padding.Empty };
            var title = new Label
            {
                Text = "Facturas",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(S(20), S(12))
            };
            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, S(11), S(14), 0),
                BackColor = Navy
            };
            var lbl = new Label
            {
                Text = "Moneda",
                ForeColor = Color.FromArgb(0xd6, 0xde, 0xf0),
                AutoSize = true,
                Margin = new Padding(0, S(6), S(6), 0)
            };
            cbSymbol = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = S(150),
                Margin = new Padding(0, S(3), S(16), 0)
            };
            cbSymbol.Items.AddRange(new object[] { "L · Lempiras", "Q · Quetzales", "$ · Dólares", "Sin símbolo" });
            btnCsv = GhostBtn("Exportar CSV");
            btnBackup = GhostBtn("Respaldo");
            btnImport = GhostBtn("Importar");
            flow.Controls.Add(lbl);
            flow.Controls.Add(cbSymbol);
            flow.Controls.Add(btnCsv);
            flow.Controls.Add(btnBackup);
            flow.Controls.Add(btnImport);
            p.Controls.Add(flow);
            p.Controls.Add(title);
            return p;
        }

        private Control BuildFolderBar()
        {
            var p = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(0xe6, 0xf4, 0xe7),
                WrapContents = false,
                Padding = new Padding(S(20), S(5), S(10), 0),
                Margin = Padding.Empty
            };
            lblFolder = new Label { AutoSize = true, Margin = new Padding(0, S(6), S(14), 0) };
            btnFolder = MakeBtn("Cambiar carpeta", false, false);
            btnFolder.MinimumSize = new Size(0, S(28));
            btnOpenFolder = MakeBtn("Abrir carpeta", false, false);
            btnOpenFolder.MinimumSize = new Size(0, S(28));
            p.Controls.Add(lblFolder);
            p.Controls.Add(btnFolder);
            p.Controls.Add(btnOpenFolder);
            lblSaved = new Label { AutoSize = true, ForeColor = Muted, Margin = new Padding(S(14), S(6), 0, 0) };
            p.Controls.Add(lblSaved);
            return p;
        }

        private Control BuildContent()
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, S(500)));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            t.Controls.Add(BuildRegister(), 0, 0);
            t.Controls.Add(BuildInvoicePanel(), 1, 0);
            return t;
        }

        // ---------- registro (izquierda) ----------
        private Label StatCell(TableLayoutPanel t, int col, string title, Color valueColor)
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = BgC, Margin = new Padding(0, 0, S(8), S(8)), Padding = new Padding(S(8), S(6), S(6), S(4)) };
            var v = new Label
            {
                Dock = DockStyle.Fill,
                Text = "-",
                ForeColor = valueColor,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            var l = new Label
            {
                Dock = DockStyle.Top,
                Height = S(18),
                Text = title.ToUpperInvariant(),
                ForeColor = Muted,
                Font = new Font("Segoe UI", 7.5f)
            };
            p.Controls.Add(v);
            p.Controls.Add(l);
            t.Controls.Add(p, col, 0);
            return v;
        }

        private Control BuildRegister()
        {
            var outer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(S(16), S(14), S(8), S(14)) };
            var card = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(S(12))
            };
            card.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            card.RowStyles.Add(new RowStyle(SizeType.Absolute, S(34)));
            card.RowStyles.Add(new RowStyle(SizeType.Absolute, S(64)));
            card.RowStyles.Add(new RowStyle(SizeType.Absolute, S(40)));
            card.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            card.RowStyles.Add(new RowStyle(SizeType.Absolute, S(40)));

            card.Controls.Add(new Label
            {
                Text = "Registro global",
                Dock = DockStyle.Fill,
                ForeColor = Navy,
                Font = new Font("Segoe UI", 11.5f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            var stats = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Margin = Padding.Empty };
            for (int i = 0; i < 3; i++) stats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            stats.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            sTotal = StatCell(stats, 0, "Facturado", Ink);
            sPaid = StatCell(stats, 1, "Cobrado", OkC);
            sPend = StatCell(stats, 2, "Pendiente", WarnC);
            card.Controls.Add(stats, 0, 1);

            var filters = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, S(120)));
            filters.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tbSearch = new TextBox { Dock = DockStyle.Top, Margin = new Padding(0, S(4), S(8), 0) };
            SetCue(tbSearch, "Buscar cliente o No.");
            cbFilter = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, S(4), 0, 0) };
            cbFilter.Items.AddRange(new object[] { "Todas", "Pendientes", "Pagadas" });
            cbFilter.SelectedIndex = 0;
            filters.Controls.Add(tbSearch, 0, 0);
            filters.Controls.Add(cbFilter, 1, 0);
            card.Controls.Add(filters, 0, 2);

            lv = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            lv.Columns.Add("No.", S(54));
            lv.Columns.Add("Cliente", S(140));
            lv.Columns.Add("Fecha", S(86));
            lv.Columns.Add("Total", S(100), HorizontalAlignment.Right);
            lv.Columns.Add("Estado", S(76));
            var cms = new ContextMenuStrip();
            cms.Items.Add("Cambiar estado (Pendiente / Pagada)", null, (s, e) => ToggleSelected());
            cms.Items.Add("Duplicar", null, (s, e) => DuplicateInvoice());
            cms.Items.Add("Eliminar", null, (s, e) => DeleteInvoice());
            lv.ContextMenuStrip = cms;
            card.Controls.Add(lv, 0, 3);

            card.Controls.Add(new Label
            {
                Text = "Selecciona una factura para abrirla. Doble clic (o clic derecho) para cambiar su estado.",
                Dock = DockStyle.Fill,
                ForeColor = Muted,
                Font = new Font("Segoe UI", 8.5f)
            }, 0, 4);

            outer.Controls.Add(card);
            return outer;
        }

        private void SetCue(TextBox t, string text)
        {
            EventHandler apply = (s, e) => SendMessage(t.Handle, 0x1501, (IntPtr)1, text);
            if (t.IsHandleCreated) apply(null, null);
            else t.HandleCreated += apply;
        }

        // ---------- factura (derecha) ----------
        private Control BuildInvoicePanel()
        {
            var outer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(S(8), S(14), S(16), S(14)) };
            var sheet = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(S(20), S(12), S(20), S(12)) };

            // campos
            tbNumber = new TextBox();
            dtDate = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy" };
            cbStatus = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            cbStatus.Items.AddRange(new object[] { "Pendiente", "Pagada" });
            dtPaid = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy" };
            tbClient = new TextBox();
            tbTax = new TextBox();
            tbAddress = new TextBox();
            fldPaid = Field("Fecha de pago", dtPaid);

            var row1 = Row(54,Field("Factura No.", tbNumber), Field("Fecha", dtDate), Field("Estado", cbStatus), fldPaid);
            var row2 = Row(54,Field("Cliente", tbClient), Field("RTN / NIT", tbTax), Field("Dirección", tbAddress));

            // cabecera
            var head = new Panel { Dock = DockStyle.Top, Height = S(50) };
            var line = new Panel { Dock = DockStyle.Bottom, Height = S(3), BackColor = Navy };
            lblHdrNo = new Label
            {
                Dock = DockStyle.Right,
                Width = S(260),
                TextAlign = ContentAlignment.BottomRight,
                ForeColor = Muted,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold)
            };
            var big = new Label
            {
                Text = "FACTURA",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft,
                ForeColor = Navy,
                Font = new Font("Segoe UI", 22f, FontStyle.Bold)
            };
            head.Controls.Add(big);
            head.Controls.Add(lblHdrNo);
            head.Controls.Add(line);

            // tabla de líneas
            grid = BuildGrid();

            var lineBtns = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = S(42),Padding = new Padding(0, S(6), 0, 0) };
            btnAddLine = MakeBtn("+ Agregar línea", false, false);
            btnDelLine = MakeBtn("Quitar línea seleccionada", false, false);
            lineBtns.Controls.Add(btnAddLine);
            lineBtns.Controls.Add(btnDelLine);

            // notas y totales
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = S(122) };
            tbNotes = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
            var notesPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 0, S(16), 0) };
            notesPanel.Controls.Add(tbNotes);
            notesPanel.Controls.Add(new Label
            {
                Text = "NOTAS / CONDICIONES DE PAGO",
                Dock = DockStyle.Top,
                Height = S(22),
                ForeColor = Muted,
                Font = new Font("Segoe UI", 8f, FontStyle.Bold)
            });

            var totals = new TableLayoutPanel { Dock = DockStyle.Right, Width = S(340), ColumnCount = 2, RowCount = 4 };
            totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            for (int i = 0; i < 3; i++) totals.RowStyles.Add(new RowStyle(SizeType.Absolute, S(26)));
            totals.RowStyles.Add(new RowStyle(SizeType.Absolute, S(42)));

            Label TL(string t) { return new Label { Text = t, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, ForeColor = Muted, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) }; }
            Label TV() { return new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI", 10f, FontStyle.Bold), Padding = new Padding(0, 0, S(8), 0) }; }

            lblSub = TV();
            lblDisc = TV();
            tbDiscount = new TextBox { Width = S(60), TextAlign = HorizontalAlignment.Right, BackColor = Cell };
            var pct = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Padding = new Padding(0, S(1), S(8), 0) };
            pct.Controls.Add(new Label { Text = "%", AutoSize = true, Margin = new Padding(S(2), S(3), 0, 0) });
            pct.Controls.Add(tbDiscount);

            totals.Controls.Add(TL("Subtotal"), 0, 0);
            totals.Controls.Add(lblSub, 1, 0);
            totals.Controls.Add(TL("Descuento"), 0, 1);
            totals.Controls.Add(pct, 1, 1);
            totals.Controls.Add(TL("Valor del descuento"), 0, 2);
            totals.Controls.Add(lblDisc, 1, 2);
            var totalL = new Label { Text = "TOTAL", Dock = DockStyle.Fill, BackColor = Navy, ForeColor = Color.White, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI", 12f, FontStyle.Bold), Padding = new Padding(0, 0, S(8), 0) };
            lblTotal = new Label { Dock = DockStyle.Fill, BackColor = Navy, ForeColor = Color.White, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI", 13f, FontStyle.Bold), Padding = new Padding(0, 0, S(8), 0) };
            totals.Controls.Add(totalL, 0, 3);
            totals.Controls.Add(lblTotal, 1, 3);

            bottom.Controls.Add(notesPanel);
            bottom.Controls.Add(totals);

            // orden de acople: lo último que se agrega queda más afuera
            sheet.Controls.Add(grid);
            sheet.Controls.Add(lineBtns);
            sheet.Controls.Add(bottom);
            sheet.Controls.Add(row2);
            sheet.Controls.Add(row1);
            sheet.Controls.Add(head);

            // barra de herramientas
            var toolbar = new Panel { Dock = DockStyle.Top, Height = S(46) };
            var tools = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(0, S(4), 0, 0) };
            btnNew = MakeBtn("+ Nueva factura", true, false);
            btnDup = MakeBtn("Duplicar", false, false);
            btnPrint = MakeBtn("Imprimir / PDF", false, false);
            btnCopyImg = MakeBtn("Copiar imagen", false, false);
            btnSaveImg = MakeBtn("Guardar imagen", false, false);
            btnDel = MakeBtn("Eliminar", false, true);
            tools.Controls.AddRange(new Control[] { btnNew, btnDup, btnPrint, btnCopyImg, btnSaveImg, btnDel });
            toolbar.Controls.Add(tools);

            outer.Controls.Add(sheet);
            outer.Controls.Add(toolbar);
            return outer;
        }

        private DataGridView BuildGrid()
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToOrderColumns = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                EnableHeadersVisualStyles = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = S(32),
                GridColor = LineC,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };
            g.RowTemplate.Height = S(28);
            g.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Navy,
                ForeColor = Color.White,
                SelectionBackColor = Navy,
                SelectionForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            };

            var right = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Cell };
            var rightRo = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            var edit = new DataGridViewCellStyle { BackColor = Cell };

            g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "#", Width = S(40), ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = new DataGridViewCellStyle { ForeColor = Muted, Alignment = DataGridViewContentAlignment.MiddleCenter } });
            g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Descripción", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = edit });
            g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Cantidad", Width = S(100), SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = right });
            g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Precio unitario", Width = S(130), SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = right });
            g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Subtotal", Width = S(140), ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = rightRo });
            return g;
        }

        // =====================================================================
        //  Datos
        // =====================================================================
        private Invoice Find(string id)
        {
            return db.Invoices.FirstOrDefault(i => i.Id == id);
        }

        private void RemoveInvoice(string id)
        {
            db.Invoices.RemoveAll(i => i.Id == id);
            if (!db.Deleted.Contains(id)) db.Deleted.Add(id);
        }

        private void EnsureCurrent()
        {
            cur = string.IsNullOrEmpty(db.Current) ? null : Find(db.Current);
            if (cur != null && U.IsBlank(cur) && db.Invoices.Any(i => !U.IsBlank(i)))
            {
                RemoveInvoice(cur.Id);
                cur = null;
            }
            if (cur == null)
            {
                if (db.Invoices.Count == 0)
                {
                    cur = U.NewInvoice(db);
                    db.Invoices.Add(cur);
                }
                else cur = U.Sorted(db.Invoices)[0];
            }
            db.Current = cur.Id;
        }

        private void StartData()
        {
            folder = Storage.ReadFolder();
            try { Directory.CreateDirectory(folder); }
            catch
            {
                folder = Storage.FallbackFolder();
                Directory.CreateDirectory(folder);
            }

            string notice;
            db = Storage.Load(folder, out notice) ?? new Db();
            startNotice = notice;

            foreach (var i in db.Invoices.Where(i => i.Id != db.Current && U.IsBlank(i)).ToList())
                RemoveInvoice(i.Id);
            EnsureCurrent();
        }

        private void ScheduleSave()
        {
            saveTimer.Stop();
            saveTimer.Start();
        }

        private void CommitEdits()
        {
            try { grid.EndEdit(); } catch { }
        }

        private void SaveNow(bool commit)
        {
            saveTimer.Stop();
            if (commit) CommitEdits();
            try
            {
                db.Current = cur != null ? cur.Id : null;
                Storage.Save(folder, db);
                lblSaved.ForeColor = Muted;
                lblSaved.Text = "Guardado a las " + DateTime.Now.ToString("HH:mm");
                errShown = false;
            }
            catch (Exception ex)
            {
                lblSaved.ForeColor = Color.FromArgb(0xb4, 0x23, 0x18);
                lblSaved.Text = "No se pudo guardar";
                if (!errShown)
                {
                    errShown = true;
                    MessageBox.Show(this, "No pude guardar en la carpeta:\n" + folder + "\n\n" + ex.Message +
                        "\n\nUsa «Cambiar carpeta» para elegir otra, o «Respaldo» para guardar una copia.",
                        "Facturas", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        // =====================================================================
        //  Eventos
        // =====================================================================
        private void HookEvents()
        {
            tbNumber.TextChanged += (s, e) => EditField(i => i.Number = tbNumber.Text);
            tbClient.TextChanged += (s, e) => EditField(i => i.Client = tbClient.Text);
            tbTax.TextChanged += (s, e) => EditField(i => i.TaxId = tbTax.Text);
            tbAddress.TextChanged += (s, e) => EditField(i => i.Address = tbAddress.Text);
            tbNotes.TextChanged += (s, e) => EditField(i => i.Notes = tbNotes.Text);
            tbDiscount.TextChanged += (s, e) => EditField(i => i.Discount = tbDiscount.Text);
            dtDate.ValueChanged += (s, e) => EditField(i => i.Date = dtDate.Value.ToString("yyyy-MM-dd", U.Inv));
            dtPaid.ValueChanged += (s, e) => EditField(i => i.PaidDate = dtPaid.Value.ToString("yyyy-MM-dd", U.Inv));
            cbStatus.SelectedIndexChanged += (s, e) => OnStatusChanged();

            cbSymbol.SelectedIndexChanged += (s, e) =>
            {
                if (loading) return;
                db.Symbol = Symbols[Math.Max(0, cbSymbol.SelectedIndex)];
                UpdateTotals();
                RefreshRegister();
                ScheduleSave();
            };

            tbSearch.TextChanged += (s, e) => RefreshRegister();
            cbFilter.SelectedIndexChanged += (s, e) => RefreshRegister();

            lv.SelectedIndexChanged += (s, e) =>
            {
                if (regLoading || lv.SelectedItems.Count == 0) return;
                OpenInvoice((string)lv.SelectedItems[0].Tag);
            };
            lv.DoubleClick += (s, e) => ToggleSelected();

            grid.CellEndEdit += OnGridCellEndEdit;
            btnAddLine.Click += (s, e) => AddLine();
            btnDelLine.Click += (s, e) => RemoveLine();

            btnNew.Click += (s, e) => NewInvoice();
            btnDup.Click += (s, e) => DuplicateInvoice();
            btnDel.Click += (s, e) => DeleteInvoice();
            btnPrint.Click += (s, e) => PrintInvoice();
            btnCopyImg.Click += (s, e) => CopyImage();
            btnSaveImg.Click += (s, e) => SaveImage();

            btnCsv.Click += (s, e) => ExportCsv();
            btnBackup.Click += (s, e) => Backup();
            btnImport.Click += (s, e) => ImportBackup();
            btnFolder.Click += (s, e) => ChangeFolder();
            btnOpenFolder.Click += (s, e) =>
            {
                try { Process.Start("explorer.exe", "\"" + folder + "\""); } catch { }
            };
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.N)) { NewInvoice(); return true; }
            if (keyData == (Keys.Control | Keys.P)) { PrintInvoice(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void EditField(Action<Invoice> apply)
        {
            if (loading || cur == null) return;
            apply(cur);
            AfterEdit();
        }

        private void AfterEdit()
        {
            if (cur != null) cur.Updated = U.NowMs();
            UpdateTotals();
            RefreshRegister();
            ScheduleSave();
        }

        private void OnStatusChanged()
        {
            if (loading || cur == null || cbStatus.SelectedItem == null) return;
            cur.Status = cbStatus.SelectedItem.ToString();
            if (cur.Status == "Pagada")
            {
                if (string.IsNullOrEmpty(cur.PaidDate)) cur.PaidDate = U.Today();
                loading = true;
                dtPaid.Value = U.ParseDate(cur.PaidDate);
                loading = false;
            }
            else cur.PaidDate = "";
            fldPaid.Visible = cur.Status == "Pagada";
            AfterEdit();
        }

        private void OnGridCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (loading || cur == null || e.RowIndex < 0 || e.RowIndex >= cur.Items.Count) return;
            if (e.ColumnIndex < 1 || e.ColumnIndex > 3) return;
            string v = Convert.ToString(grid[e.ColumnIndex, e.RowIndex].Value) ?? "";
            var it = cur.Items[e.RowIndex];
            if (e.ColumnIndex == 1) it.Desc = v;
            else if (e.ColumnIndex == 2) it.Qty = v;
            else it.Price = v;
            AfterEdit();
        }

        // =====================================================================
        //  Mostrar datos
        // =====================================================================
        private void RefreshAll()
        {
            loading = true;
            int idx = Array.IndexOf(Symbols, db.Symbol);
            cbSymbol.SelectedIndex = idx >= 0 ? idx : 0;
            loading = false;
            LoadInvoice();
            RefreshRegister();
            UpdateFolderBar();
        }

        private void UpdateFolderBar()
        {
            lblFolder.Text = "Guardando en la carpeta: " + folder;
        }

        private void LoadInvoice()
        {
            if (cur == null) return;
            loading = true;
            tbNumber.Text = cur.Number;
            dtDate.Value = U.ParseDate(cur.Date);
            cbStatus.SelectedItem = cur.Status;
            dtPaid.Value = U.ParseDate(cur.PaidDate);
            fldPaid.Visible = cur.Status == "Pagada";
            tbClient.Text = cur.Client;
            tbTax.Text = cur.TaxId;
            tbAddress.Text = cur.Address;
            tbNotes.Text = cur.Notes;
            tbDiscount.Text = cur.Discount;
            LoadItems();
            loading = false;
            UpdateTotals();
        }

        private void LoadItems()
        {
            bool prev = loading;
            loading = true;
            grid.Rows.Clear();
            for (int i = 0; i < cur.Items.Count; i++)
            {
                var it = cur.Items[i];
                grid.Rows.Add(i + 1, it.Desc, it.Qty, it.Price, "");
            }
            loading = prev;
        }

        private void UpdateTotals()
        {
            if (cur == null) return;
            string sym = db.Symbol;
            for (int i = 0; i < grid.Rows.Count && i < cur.Items.Count; i++)
            {
                double sub = U.Num(cur.Items[i].Qty) * U.Num(cur.Items[i].Price);
                grid[4, i].Value = sub > 0 ? U.Money(sub, sym) : "";
            }
            Totals t = U.Calc(cur);
            lblSub.Text = U.Money(t.Sub, sym);
            lblDisc.Text = t.Disc != 0 ? "- " + U.Money(t.Disc, sym) : U.Money(0, sym);
            lblTotal.Text = U.Money(t.Total, sym);
            lblHdrNo.Text = string.IsNullOrEmpty(cur.Number) ? "" : "No. " + cur.Number;
        }

        private void RefreshRegister()
        {
            string sym = db.Symbol;
            double total = 0, paid = 0;
            foreach (var i in db.Invoices)
            {
                if (U.IsBlank(i)) continue;
                double t = U.R2(U.Calc(i).Total);
                total += t;
                if (i.Status == "Pagada") paid += t;
            }
            sTotal.Text = U.Money(total, sym);
            sPaid.Text = U.Money(paid, sym);
            sPend.Text = U.Money(total - paid, sym);

            string q = tbSearch.Text.Trim().ToLowerInvariant();
            int fs = cbFilter.SelectedIndex;
            var list = U.Sorted(db.Invoices).Where(i =>
            {
                if (fs == 1 && i.Status != "Pendiente") return false;
                if (fs == 2 && i.Status != "Pagada") return false;
                if (q.Length > 0 &&
                    (i.Client ?? "").ToLowerInvariant().IndexOf(q, StringComparison.Ordinal) < 0 &&
                    (i.Number ?? "").ToLowerInvariant().IndexOf(q, StringComparison.Ordinal) < 0) return false;
                return true;
            }).ToList();

            regLoading = true;
            lv.BeginUpdate();
            lv.Items.Clear();
            foreach (var i in list)
            {
                bool ok = i.Status == "Pagada";
                var item = new ListViewItem(i.Number ?? "") { Tag = i.Id, UseItemStyleForSubItems = false };
                var cli = item.SubItems.Add(string.IsNullOrEmpty(i.Client) ? "(sin cliente)" : i.Client);
                if (string.IsNullOrEmpty(i.Client)) cli.ForeColor = Muted;
                item.SubItems.Add(U.FmtDate(i.Date));
                item.SubItems.Add(U.Money(U.Calc(i).Total, sym));
                var st = item.SubItems.Add(i.Status);
                st.ForeColor = ok ? OkC : WarnC;
                st.Font = statusFont;
                lv.Items.Add(item);
                if (i.Id == db.Current) item.Selected = true;
            }
            lv.EndUpdate();
            regLoading = false;
        }

        // =====================================================================
        //  Acciones
        // =====================================================================
        private void OpenInvoice(string id)
        {
            CommitEdits();
            if (cur != null && cur.Id != id && U.IsBlank(cur)) RemoveInvoice(cur.Id);
            var inv = Find(id);
            if (inv == null) return;
            cur = inv;
            db.Current = id;
            LoadInvoice();
            RefreshRegister();
            ScheduleSave();
        }

        private void NewInvoice()
        {
            CommitEdits();
            if (cur != null && U.IsBlank(cur))
            {
                tbClient.Focus();
                return;
            }
            var inv = U.NewInvoice(db);
            db.Invoices.Add(inv);
            cur = inv;
            db.Current = inv.Id;
            LoadInvoice();
            RefreshRegister();
            ScheduleSave();
            tbClient.Focus();
        }

        private void DuplicateInvoice()
        {
            CommitEdits();
            if (cur == null) return;
            var copy = new Invoice
            {
                Id = U.Uid(),
                Number = U.NextNumber(db),
                Date = U.Today(),
                Client = cur.Client,
                TaxId = cur.TaxId,
                Address = cur.Address,
                Discount = cur.Discount,
                Notes = cur.Notes,
                Status = "Pendiente",
                PaidDate = "",
                Updated = U.NowMs()
            };
            foreach (var it in cur.Items) copy.Items.Add(new Item { Desc = it.Desc, Qty = it.Qty, Price = it.Price });
            db.Invoices.Add(copy);
            cur = copy;
            db.Current = copy.Id;
            LoadInvoice();
            RefreshRegister();
            ScheduleSave();
        }

        private void DeleteInvoice()
        {
            CommitEdits();
            if (cur == null) return;
            string label = "No. " + cur.Number + (string.IsNullOrEmpty(cur.Client) ? "" : " (" + cur.Client + ")");
            if (MessageBox.Show(this, "¿Eliminar la factura " + label + "?\nEsta acción no se puede deshacer.",
                    "Eliminar factura", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            RemoveInvoice(cur.Id);
            var rest = U.Sorted(db.Invoices);
            if (rest.Count > 0) cur = rest[0];
            else
            {
                cur = U.NewInvoice(db);
                db.Invoices.Add(cur);
            }
            db.Current = cur.Id;
            LoadInvoice();
            RefreshRegister();
            ScheduleSave();
        }

        private void ToggleSelected()
        {
            if (lv.SelectedItems.Count == 0) return;
            var inv = Find((string)lv.SelectedItems[0].Tag);
            if (inv == null) return;
            CommitEdits();
            if (inv.Status == "Pagada")
            {
                inv.Status = "Pendiente";
                inv.PaidDate = "";
            }
            else
            {
                inv.Status = "Pagada";
                if (string.IsNullOrEmpty(inv.PaidDate)) inv.PaidDate = U.Today();
            }
            inv.Updated = U.NowMs();
            if (inv.Id == db.Current) LoadInvoice();
            RefreshRegister();
            ScheduleSave();
        }

        private void AddLine()
        {
            if (cur == null) return;
            CommitEdits();
            cur.Items.Add(U.BlankItem());
            cur.Updated = U.NowMs();
            LoadItems();
            UpdateTotals();
            ScheduleSave();
            int last = grid.Rows.Count - 1;
            if (last >= 0)
            {
                grid.CurrentCell = grid[1, last];
                grid.BeginEdit(true);
            }
        }

        private void RemoveLine()
        {
            if (cur == null || grid.CurrentCell == null) return;
            int i = grid.CurrentCell.RowIndex;
            CommitEdits();
            if (i < 0 || i >= cur.Items.Count) return;
            cur.Items.RemoveAt(i);
            if (cur.Items.Count == 0) cur.Items.Add(U.BlankItem());
            cur.Updated = U.NowMs();
            LoadItems();
            UpdateTotals();
            RefreshRegister();
            ScheduleSave();
        }

        // ---------- imprimir / imagen ----------
        private void PrintInvoice()
        {
            CommitEdits();
            if (cur == null) return;
            using (var bmp = Render.ToBitmap(cur, db.Symbol, 3))
            using (var pd = new PrintDocument())
            {
                pd.DocumentName = "Factura " + cur.Number;
                pd.PrintPage += (s, e) =>
                {
                    float sc = Math.Min((float)e.MarginBounds.Width / bmp.Width, (float)e.MarginBounds.Height / bmp.Height);
                    e.Graphics.DrawImage(bmp, e.MarginBounds.Left, e.MarginBounds.Top, bmp.Width * sc, bmp.Height * sc);
                    e.HasMorePages = false;
                };
                using (var dlg = new PrintDialog { Document = pd, UseEXDialog = true })
                {
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
                    try { pd.Print(); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "No se pudo imprimir: " + ex.Message, "Facturas",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private string ImageName()
        {
            string name = "Factura_" + (cur.Number ?? "") + (string.IsNullOrEmpty(cur.Client) ? "" : "_" + cur.Client);
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name.Replace(' ', '_');
        }

        private void CopyImage()
        {
            CommitEdits();
            if (cur == null) return;
            try
            {
                using (var bmp = Render.ToBitmap(cur, db.Symbol, 2))
                    Clipboard.SetImage(bmp);
                lblSaved.ForeColor = OkC;
                lblSaved.Text = "Imagen copiada: pégala con Ctrl+V";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No pude copiar la imagen: " + ex.Message + "\nUsa «Guardar imagen».", "Facturas",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SaveImage()
        {
            CommitEdits();
            if (cur == null) return;
            using (var dlg = new SaveFileDialog { Filter = "Imagen PNG (*.png)|*.png", FileName = ImageName() + ".png" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    using (var bmp = Render.ToBitmap(cur, db.Symbol, 2))
                        bmp.Save(dlg.FileName, System.Drawing.Imaging.ImageFormat.Png);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No pude guardar la imagen: " + ex.Message, "Facturas",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        // ---------- CSV / respaldo ----------
        private static string CsvCell(string v)
        {
            v = v ?? "";
            return v.IndexOfAny(new[] { '"', ',', '\n', ';', '\r' }) >= 0 ? "\"" + v.Replace("\"", "\"\"") + "\"" : v;
        }

        private void ExportCsv()
        {
            CommitEdits();
            var rows = U.Sorted(db.Invoices).Where(i => !U.IsBlank(i)).ToList();
            if (rows.Count == 0)
            {
                MessageBox.Show(this, "Todavía no hay facturas con datos para exportar.", "Facturas",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (var dlg = new SaveFileDialog { Filter = "CSV para Excel (*.csv)|*.csv", FileName = "registro_facturas_" + U.Today() + ".csv" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                var sb = new StringBuilder();
                sb.Append("Número,Fecha,Cliente,RTN/NIT,Subtotal,Descuento,Total,Estado,Fecha de pago\r\n");
                foreach (var i in rows)
                {
                    Totals t = U.Calc(i);
                    var cells = new[]
                    {
                        i.Number, i.Date, i.Client, i.TaxId,
                        U.R2(t.Sub).ToString("0.00", U.Inv), U.R2(t.Disc).ToString("0.00", U.Inv),
                        U.R2(t.Total).ToString("0.00", U.Inv), i.Status, i.PaidDate
                    };
                    sb.Append(string.Join(",", cells.Select(CsvCell))).Append("\r\n");
                }
                try { File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true)); }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No pude guardar el archivo: " + ex.Message, "Facturas",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void Backup()
        {
            CommitEdits();
            using (var dlg = new SaveFileDialog { Filter = "Respaldo (*.json)|*.json", FileName = "respaldo_facturas_" + U.Today() + ".json" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    db.Current = cur != null ? cur.Id : null;
                    File.WriteAllText(dlg.FileName, Storage.ToJson(db), new UTF8Encoding(false));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No pude guardar el respaldo: " + ex.Message, "Facturas",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void ImportBackup()
        {
            CommitEdits();
            using (var dlg = new OpenFileDialog { Filter = "Respaldo (*.json)|*.json|Todos los archivos|*.*" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                Db d;
                try { d = Storage.FromJson(File.ReadAllText(dlg.FileName, Encoding.UTF8)); }
                catch
                {
                    MessageBox.Show(this, "No pude leer ese archivo. Elige un respaldo guardado desde Facturas (o desde la versión web).",
                        "Facturas", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (MessageBox.Show(this, "Esto reemplazará las facturas que tienes ahora por las del respaldo (" + d.Invoices.Count + ").\n¿Continuar?",
                        "Importar respaldo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

                var oldIds = db.Invoices.Select(i => i.Id).ToList();
                foreach (var i in d.Invoices) i.Updated = U.NowMs();
                db.Invoices = d.Invoices;
                var newIds = new HashSet<string>(db.Invoices.Select(i => i.Id));
                db.Deleted = oldIds.Where(id => !newIds.Contains(id)).ToList();
                db.Symbol = d.Symbol;
                db.Current = d.Current;
                EnsureCurrent();
                RefreshAll();
                SaveNow(false);
            }
        }

        private void ChangeFolder()
        {
            CommitEdits();
            using (var dlg = new FolderBrowserDialog { Description = "Elige la carpeta donde se guardarán tus facturas", SelectedPath = folder })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                string nf = dlg.SelectedPath;
                try { Directory.CreateDirectory(nf); }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No pude usar esa carpeta: " + ex.Message, "Facturas",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string notice;
                Db other = Storage.Load(nf, out notice);
                folder = nf;
                Storage.SaveFolder(nf);
                if (other != null) U.Merge(db, other);
                EnsureCurrent();
                RefreshAll();
                SaveNow(false);
                if (notice != null)
                    MessageBox.Show(this, notice, "Facturas", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
