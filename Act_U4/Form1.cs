using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace Act_U4
{
    public partial class Form1 : Form
    {
        // Modelos de datos
        private CsvData? _csvData;
        private Dictionary<string, double>? _barData;
        private Dictionary<string, double>? _pieData;

        // Controles de UI
        private Button _btnLoad = null!;
        private Button _btnExportPdf = null!;
        private TabControl _tabControl = null!;
        private DataGridView _gridData = null!;
        private DataGridView _gridHierarchy = null!;

        // Paneles para las gráficas
        private Panel _panelBarChart = null!;
        private Panel _panelPieChart = null!;

        // Colores temáticos de tecnología
        private Color[] techColors = {
            Color.FromArgb(0, 90, 156),   // Azul oscuro
            Color.FromArgb(0, 150, 214),  // Azul claro
            Color.FromArgb(0, 191, 178),  // Turquesa
            Color.FromArgb(108, 117, 125),// Gris
            Color.FromArgb(52, 58, 64),   // Gris oscuro
            Color.FromArgb(23, 162, 184)  // Info Cyan
        };
        public Form1()
        {
            InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            ConfigurarInterfazManual();

        }
        private void ConfigurarInterfazManual()
        {
            this.Text = "Sistema de Reportes - Productos vendidos ";
            this.Size = new Size(1200, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9F);
            this.BackColor = Color.FromArgb(244, 246, 249); // Gris muy claro

            // 1. Panel Superior
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(33, 37, 41) };

            _btnLoad = new Button
            {
                Text = "Cargar Datos (CSV)",
                Location = new Point(20, 12),
                Size = new Size(200, 36),
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnLoad.FlatAppearance.BorderSize = 0;
            _btnLoad.Click += BtnLoad_Click;

            _btnExportPdf = new Button
            {
                Text = "Exportar Reporte PDF",
                Location = new Point(230, 12),
                Size = new Size(200, 36),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnExportPdf.FlatAppearance.BorderSize = 0;
            _btnExportPdf.Click += BtnExportPdf_Click;

            topPanel.Controls.Add(_btnLoad);
            topPanel.Controls.Add(_btnExportPdf);

            // 2. TabControl (Pestañas)
            _tabControl = new TabControl { Dock = DockStyle.Fill, ItemSize = new Size(150, 35), Padding = new Point(15, 8) };

            // Pestaña A: Modo Tabular
            TabPage tabDatos = new TabPage("a. Modo Tabular");
            _gridData = CreateStyledGrid();
            Label lblTabularDesc = new Label
            {
                Dock = DockStyle.Top,
                Height = 60,
                Text = "CARACTERÍSTICAS DEL REPORTE TABULAR: Presenta la información estructurada en filas y columnas. Es exhaustivo, permite realizar búsquedas de datos específicos, cruzar variables de forma exacta y revisar los registros crudos de las transacciones.",
                Font = new Font("Segoe UI", 10, FontStyle.Italic),
                Padding = new Padding(10),
                BackColor = Color.White
            };
            tabDatos.Controls.Add(_gridData);
            tabDatos.Controls.Add(lblTabularDesc);

            // Pestaña B: Modo Gráfico (Barras)
            TabPage tabBarras = new TabPage("b. Gráficas (Ingresos)");
            _panelBarChart = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            _panelBarChart.Paint += DrawBarChart;
            _panelBarChart.Resize += (s, e) => _panelBarChart.Invalidate();
            tabBarras.Controls.Add(_panelBarChart);

            // Pestaña B: Modo Gráfico (Pastel)
            TabPage tabPastel = new TabPage("b. Gráficas (Volumen)");
            _panelPieChart = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            _panelPieChart.Paint += DrawPieChart;
            _panelPieChart.Resize += (s, e) => _panelPieChart.Invalidate();
            tabPastel.Controls.Add(_panelPieChart);

            // Pestaña C: Jerarquía e Importancia
            TabPage tabStats = new TabPage("c. Jerarquía e Importancia");
            _gridHierarchy = CreateStyledGrid();
            Label lblJerarquiaDesc = new Label
            {
                Dock = DockStyle.Top,
                Height = 60,
                Text = "JERARQUÍA E IMPORTANCIA: Este reporte clasifica a los elementos de mayor a menor rendimiento (Ranking). Permite distinguir rápidamente quiénes o cuáles son los activos más valiosos de la empresa, facilitando la toma de decisiones estratégicas y recompensas.",
                Font = new Font("Segoe UI", 10, FontStyle.Italic),
                Padding = new Padding(10),
                BackColor = Color.White
            };
            tabStats.Controls.Add(_gridHierarchy);
            tabStats.Controls.Add(lblJerarquiaDesc);

            _tabControl.TabPages.Add(tabDatos);
            _tabControl.TabPages.Add(tabBarras);
            _tabControl.TabPages.Add(tabPastel);
            _tabControl.TabPages.Add(tabStats);

            this.Controls.Add(_tabControl);
            this.Controls.Add(topPanel);
        }

        private DataGridView CreateStyledGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(240, 248, 255) } // AliceBlue
            };
        }

        private List<string> ParseCsvLine(string line)
        {
            List<string> result = new List<string>();
            bool inQuotes = false;
            string currentField = "";

            foreach (char c in line)
            {
                if (c == '"') inQuotes = !inQuotes;
                else if (c == ',' && !inQuotes)
                {
                    result.Add(currentField.Trim());
                    currentField = "";
                }
                else currentField += c;
            }
            result.Add(currentField.Trim());
            return result;
        }

        private void BtnLoad_Click(object? sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Archivos CSV|*.csv|Todos los archivos|*.*" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        LoadAndAnalyzeCsv(ofd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al leer el archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void LoadAndAnalyzeCsv(string filePath)
        {
            var lines = File.ReadAllLines(filePath);
            if (lines.Length < 2) throw new Exception("El archivo está vacío o no tiene datos suficientes.");

            _csvData = new CsvData();
            _csvData.Headers = ParseCsvLine(lines[0]);

            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                _csvData.Rows.Add(ParseCsvLine(lines[i]));
            }

            _gridData.Columns.Clear();
            _gridData.Rows.Clear();
            foreach (var header in _csvData.Headers) _gridData.Columns.Add(header, header);

            int rowsToShow = Math.Min(_csvData.Rows.Count, 1500);
            for (int i = 0; i < rowsToShow; i++) _gridData.Rows.Add(_csvData.Rows[i].ToArray());

            ProcesarInteligenciaDeNegocios();
        }

        private void ProcesarInteligenciaDeNegocios()
        {
            if (_csvData == null) return;

            // Esperamos columnas tipo: Producto, Categoria, Ingresos, Vendedor
            int idxProduct = _csvData.Headers.IndexOf("Producto");
            int idxCategory = _csvData.Headers.IndexOf("Categoria");
            int idxAmount = _csvData.Headers.IndexOf("Ingresos");
            int idxSalesPerson = _csvData.Headers.IndexOf("Vendedor");

            if (idxProduct < 0 || idxAmount < 0 || idxSalesPerson < 0)
            {
                MessageBox.Show("El CSV no tiene las columnas requeridas (Producto, Ingresos, Vendedor).", "Formato Incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _barData = new Dictionary<string, double>();
            _pieData = new Dictionary<string, double>();
            Dictionary<string, double> salesPersonData = new Dictionary<string, double>();

            foreach (var row in _csvData.Rows)
            {
                string product = row[idxProduct];
                string category = idxCategory >= 0 ? row[idxCategory] : "General";
                string salesPerson = row[idxSalesPerson];

                string cleanAmount = row[idxAmount].Replace("$", "").Replace(",", "").Trim();
                double.TryParse(cleanAmount, NumberStyles.Any, CultureInfo.InvariantCulture, out double amount);

                // Acumular ingresos por producto (Barras)
                if (_barData.ContainsKey(product)) _barData[product] += amount; else _barData[product] = amount;

                // Acumular ingresos por categoria (Pastel)
                if (_pieData.ContainsKey(category)) _pieData[category] += amount; else _pieData[category] = amount;

                // Acumular ingresos por vendedor (Jerarquía)
                if (salesPersonData.ContainsKey(salesPerson)) salesPersonData[salesPerson] += amount; else salesPersonData[salesPerson] = amount;
            }

            // Ordenar datos
            _barData = _barData.OrderByDescending(kv => kv.Value).Take(10).ToDictionary(kv => kv.Key, kv => kv.Value);
            salesPersonData = salesPersonData.OrderByDescending(kv => kv.Value).ToDictionary(kv => kv.Key, kv => kv.Value);

            // Llenar DataGridView de Jerarquía
            _gridHierarchy.Columns.Clear();
            _gridHierarchy.Rows.Clear();
            _gridHierarchy.Columns.Add("Pos", "Nivel de Jerarquía");
            _gridHierarchy.Columns.Add("Name", "Agente de Ventas");
            _gridHierarchy.Columns.Add("Val", "Impacto Generado (MXN)");

            int rank = 1;
            foreach (var sp in salesPersonData)
            {
                _gridHierarchy.Rows.Add($"Nivel {rank}", sp.Key, "$" + sp.Value.ToString("N2"));
                rank++;
            }

            _panelBarChart.Invalidate();
            _panelPieChart.Invalidate();
        }

        private void DrawBarChart(object? sender, PaintEventArgs e) => RenderBarChart(e.Graphics, _panelBarChart.ClientRectangle);
        private void DrawPieChart(object? sender, PaintEventArgs e) => RenderPieChart(e.Graphics, _panelPieChart.ClientRectangle);

        private void RenderBarChart(Graphics g, Rectangle rect)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.White);

            if (_barData == null || _barData.Count == 0) { DrawNoData(g, rect); return; }

            int marginX = 80, marginY = 120; // Margen inferior más amplio para las etiquetas
            int chartWidth = rect.Width - (marginX * 2);
            int chartHeight = rect.Height - (marginY * 2);
            if (chartWidth <= 0 || chartHeight <= 0) return;

            // Título y Descripción de la Rúbrica
            g.DrawString("Ingresos por Producto", new Font("Segoe UI", 16, FontStyle.Bold), Brushes.Black, new PointF(marginX, 15));
            g.DrawString("CARACTERÍSTICAS Y VENTAJAS: Reporte gráfico en barras bidimensionales.\nVentaja principal: Permite una comparación visual y rápida de magnitudes entre distintos productos.", new Font("Segoe UI", 9, FontStyle.Italic), Brushes.DimGray, new PointF(marginX, 45));

            double maxVal = _barData.Values.Max();
            double scaleY = maxVal == 0 ? 1 : chartHeight / maxVal;

            using (Pen axisPen = new Pen(Color.Black, 2))
            {
                g.DrawLine(axisPen, marginX, rect.Height - marginY, rect.Width - marginX, rect.Height - marginY);
                g.DrawLine(axisPen, marginX, 90, marginX, rect.Height - marginY);
            }

            float barTotalWidth = chartWidth / (float)_barData.Count;
            float barWidth = barTotalWidth * 0.7f;
            float spacing = barTotalWidth * 0.15f;
            int i = 0;

            foreach (var item in _barData)
            {
                float barHeight = (float)(item.Value * scaleY);
                float x = marginX + (i * barTotalWidth) + spacing;
                float y = rect.Height - marginY - barHeight;

                RectangleF barRect = new RectangleF(x, y, barWidth, barHeight);

                // Sombra
                g.FillRectangle(new SolidBrush(Color.FromArgb(20, 0, 0, 0)), x + 4, y + 4, barWidth, barHeight);

                Color c = techColors[i % techColors.Length];
                using (LinearGradientBrush brush = new LinearGradientBrush(barRect, ControlPaint.Light(c), c, LinearGradientMode.Vertical))
                    g.FillRectangle(brush, barRect);

                g.DrawRectangle(Pens.Black, x, y, barWidth, barHeight);

                // Etiqueta de valor
                string valStr = "$" + (item.Value / 1000).ToString("F1") + "k";
                SizeF valSize = g.MeasureString(valStr, this.Font);
                g.DrawString(valStr, this.Font, Brushes.Black, x + (barWidth / 2) - (valSize.Width / 2), y - 20);

                // Etiqueta de eje X (Rotada)
                string labelStr = item.Key.Length > 15 ? item.Key.Substring(0, 12) + ".." : item.Key;
                g.TranslateTransform(x + (barWidth / 2), rect.Height - marginY + 10);
                g.RotateTransform(45);
                g.DrawString(labelStr, this.Font, Brushes.Black, 0, 0);
                g.ResetTransform();
                i++;
            }
        }

        private void RenderPieChart(Graphics g, Rectangle rect)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.White);

            if (_pieData == null || _pieData.Count == 0) { DrawNoData(g, rect); return; }

            g.DrawString("Distribución de Ingresos por Categoría", new Font("Segoe UI", 16, FontStyle.Bold), Brushes.Black, new PointF(50, 15));
            g.DrawString("CARACTERÍSTICAS Y VENTAJAS: Reporte gráfico en formato de pastel o torta.\nVentaja principal: Facilita la comprensión de la participación de mercado o proporción de un " +
                "elemento frente al 100% total.", new Font("Segoe UI", 9, FontStyle.Italic), Brushes.DimGray, new PointF(50, 45));

            int diameter = Math.Min(rect.Width, rect.Height) - 180;
            if (diameter <= 0) return;

            Rectangle pieRect = new Rectangle(50, 100, diameter, diameter);
            double total = _pieData.Values.Sum();
            float startAngle = 0;
            int i = 0;

            g.FillEllipse(new SolidBrush(Color.FromArgb(15, 0, 0, 0)), pieRect.X + 5, pieRect.Y + 5, diameter, diameter);

            int legendX = pieRect.Right + 50;
            int legendY = pieRect.Y;

            foreach (var item in _pieData.OrderByDescending(x => x.Value))
            {
                float sweepAngle = (float)((item.Value / total) * 360f);
                Color c = techColors[i % techColors.Length];

                g.FillPie(new SolidBrush(c), pieRect, startAngle, sweepAngle);
                g.DrawPie(Pens.White, pieRect, startAngle, sweepAngle);

                g.FillRectangle(new SolidBrush(c), legendX, legendY, 20, 20);
                double pct = (item.Value / total) * 100;
                string legText = $"{item.Key}: ${item.Value:N0} ({pct:F1}%)";
                g.DrawString(legText, new Font("Segoe UI", 10, FontStyle.Bold), Brushes.Black, legendX + 30, legendY);

                legendY += 30;
                startAngle += sweepAngle;
                i++;
            }
        }
        private void RenderTablaDatos(Graphics g, Rectangle rect)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.White);
            if (_gridData.Rows.Count == 0) { DrawNoData(g, rect); return; }

            int marginX = 40;
            int yPos = 110;

            g.DrawString("Datos Crudos", new Font("Segoe UI", 16, FontStyle.Bold), Brushes.Black, new PointF(marginX, 15));
            g.DrawString("CARACTERÍSTICAS: Presenta la información en filas y columnas. Es exhaustivo y permite revisar cada registro exacto.", new Font("Segoe UI", 9, FontStyle.Italic), Brushes.DimGray, new PointF(marginX, 45));
            g.DrawString("*Mostrando una muestra de los primeros 18 registros.", new Font("Segoe UI", 8, FontStyle.Italic), Brushes.Gray, new PointF(marginX, 70));

            int colWidth = (rect.Width - (marginX * 2)) / _gridData.Columns.Count;

            // Encabezados
            Font headerFont = new Font("Segoe UI", 10, FontStyle.Bold);
            Rectangle headerRect = new Rectangle(marginX, yPos, rect.Width - (marginX * 2), 35);
            g.FillRectangle(new SolidBrush(Color.FromArgb(0, 123, 255)), headerRect);

            for (int c = 0; c < _gridData.Columns.Count; c++)
            {
                g.DrawString(_gridData.Columns[c].HeaderText, headerFont, Brushes.White, new PointF(marginX + (c * colWidth) + 5, yPos + 8));
            }
            yPos += 35;

            // Filas
            Font rowFont = new Font("Segoe UI", 9);
            int rowsToDraw = Math.Min(_gridData.Rows.Count, 18); // Límite para que quepa en la hoja

            for (int i = 0; i < rowsToDraw; i++)
            {
                var row = _gridData.Rows[i];
                Rectangle rowRect = new Rectangle(marginX, yPos, rect.Width - (marginX * 2), 30);

                if (i % 2 == 0) g.FillRectangle(new SolidBrush(Color.FromArgb(240, 248, 255)), rowRect);
                g.DrawRectangle(Pens.LightGray, rowRect);

                for (int c = 0; c < _gridData.Columns.Count; c++)
                {
                    string val = row.Cells[c].Value?.ToString() ?? "";
                    if (val.Length > 20) val = val.Substring(0, 17) + "...";
                    g.DrawString(val, rowFont, Brushes.Black, new PointF(marginX + (c * colWidth) + 5, yPos + 6));
                }

                yPos += 30;
                if (yPos > rect.Height - 40) break;
            }
        }
        private void RenderTablaJerarquia(Graphics g, Rectangle rect)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.White);
            if (_gridHierarchy.Rows.Count == 0) { DrawNoData(g, rect); return; }

            int marginX = 80;
            int yPos = 110;

            g.DrawString("Top Vendedores", new Font("Segoe UI", 16, FontStyle.Bold), Brushes.Black, new PointF(marginX, 15));
            g.DrawString("JERARQUÍA E IMPORTANCIA: Clasifica a los elementos de mayor a menor rendimiento.", new Font("Segoe UI", 9, FontStyle.Italic), Brushes.DimGray, new PointF(marginX, 45));

            int col1Width = 150;
            int col2Width = 350;
            int col3Width = 250;

            // Encabezados
            Font headerFont = new Font("Segoe UI", 12, FontStyle.Bold);
            Rectangle headerRect = new Rectangle(marginX, yPos, col1Width + col2Width + col3Width, 40);
            g.FillRectangle(new SolidBrush(Color.FromArgb(33, 37, 41)), headerRect);

            g.DrawString("Nivel", headerFont, Brushes.White, new PointF(marginX + 10, yPos + 10));
            g.DrawString("Agente de Ventas", headerFont, Brushes.White, new PointF(marginX + col1Width + 10, yPos + 10));
            g.DrawString("Impacto (MXN)", headerFont, Brushes.White, new PointF(marginX + col1Width + col2Width + 10, yPos + 10));

            yPos += 40;

            // Filas
            Font rowFont = new Font("Segoe UI", 11);
            for (int i = 0; i < _gridHierarchy.Rows.Count; i++)
            {
                var row = _gridHierarchy.Rows[i];
                Rectangle rowRect = new Rectangle(marginX, yPos, col1Width + col2Width + col3Width, 35);

                if (i % 2 == 0) g.FillRectangle(new SolidBrush(Color.FromArgb(240, 248, 255)), rowRect);
                g.DrawRectangle(Pens.LightGray, rowRect);

                string pos = row.Cells[0].Value?.ToString() ?? "";
                string name = row.Cells[1].Value?.ToString() ?? "";
                string val = row.Cells[2].Value?.ToString() ?? "";

                g.DrawString(pos, rowFont, Brushes.Black, new PointF(marginX + 10, yPos + 8));
                g.DrawString(name, rowFont, Brushes.Black, new PointF(marginX + col1Width + 10, yPos + 8));
                g.DrawString(val, rowFont, Brushes.Black, new PointF(marginX + col1Width + col2Width + 10, yPos + 8));

                yPos += 35;
                if (yPos > rect.Height - 50) break;
            }
        }

        private void DrawNoData(Graphics g, Rectangle rect)
        {
            g.DrawString("Cargue el archivo CSV para procesar los reportes.", new Font("Segoe UI", 12), Brushes.Gray, new PointF(rect.X + 50, rect.Y + 50));
        }

        private void BtnExportPdf_Click(object? sender, EventArgs e)
        {
            if (_barData == null || _barData.Count == 0)
            {
                MessageBox.Show("Primero carga un archivo CSV para generar el reporte.", "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Archivos PDF (*.pdf)|*.pdf", FileName = "Reporte_Sistemas_Integrales.pdf" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        GenerarPdfDocumento(sfd.FileName);
                        MessageBox.Show("¡Reporte PDF generado exitosamente!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al generar el PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void GenerarPdfDocumento(string filePath)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Reporte Ejecutivo";

            //Action<Graphics, Rectangle>[] metodosDeRenderizado = { RenderBarChart, RenderPieChart, RenderTablaDatos };
            //string[] titulosDePagina = { "a.Reporte Tabular ,b. Reporte Gráfico (Análisis Comparativo)", "b. Reporte Gráfico (Análisis Proporcional)" };
            Action<Graphics, Rectangle>[] metodosDeRenderizado = {
    RenderTablaDatos,
    RenderBarChart,
    RenderPieChart,
    RenderTablaJerarquia
};

            // Y asegúrate de que tienes estos 4 títulos (aquí es donde suele faltar código):
            string[] titulosDePagina = {
    "a. Reporte Tabular",
    "b. Reporte Gráfico (Análisis Comparativo)",
    "b. Reporte Gráfico (Análisis Proporcional)",
    "c. Jerarquía e Importancia"
};
            List<MemoryStream> streamsGuardados = new List<MemoryStream>();

            try
            {
                for (int i = 0; i < metodosDeRenderizado.Length; i++)
                {
                    PdfPage page = document.AddPage();
                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                    {
                        int imgWidth = 1000;
                        int imgHeight = 800;

                        using (Bitmap bmpTemp = new Bitmap(imgWidth, imgHeight))
                        {
                            using (Graphics gBitmap = Graphics.FromImage(bmpTemp))
                            {
                                gBitmap.Clear(Color.White);
                                using (Font fontTitulo = new Font("Segoe UI", 24, FontStyle.Bold))
                                {
                                    SizeF textSize = gBitmap.MeasureString(titulosDePagina[i], fontTitulo);
                                    gBitmap.DrawString(titulosDePagina[i], fontTitulo, Brushes.DarkBlue, new PointF((imgWidth - textSize.Width) / 2, 20));
                                }
                                metodosDeRenderizado[i](gBitmap, new Rectangle(0, 100, imgWidth, imgHeight - 100));
                            }

                            MemoryStream ms = new MemoryStream();
                            bmpTemp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            ms.Position = 0;
                            streamsGuardados.Add(ms);

                            using (XImage imagenParaPdf = XImage.FromStream(ms))
                            {
                                double margenLateral = 20;
                                double anchoDisponible = page.Width - (margenLateral * 2);
                                double escala = anchoDisponible / bmpTemp.Width;
                                gfx.DrawImage(imagenParaPdf, margenLateral, 40, anchoDisponible, bmpTemp.Height * escala);
                            }
                        }
                    }
                }
                document.Save(filePath);
            }
            finally
            {
                foreach (var ms in streamsGuardados) ms.Dispose();
            }
        }
    }

    public class CsvData
    {
        public List<string> Headers { get; set; } = new List<string>();
        public List<List<string>> Rows { get; set; } = new List<List<string>>();
    }

}

