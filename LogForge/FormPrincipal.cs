using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace LogForge
{
    public partial class FormPrincipal : Form
    {
        private List<string> logsDisponiveis = new List<string> { "Application", "Security", "System", "Setup", "ForwardedEvents" };
        private CheckedListBox chklstLogs;
        private DateTimePicker dtpInicio, dtpFim;
        private Button btnGerar, btnAbrirPasta, btnSelecionarPasta;
        private ProgressBar progressBar;
        private RadioButton rbtnTXT, rbtnCSV;
        private TextBox txtLogStatus;
        private DataGridView dgvPreview;
        private string pastaDestino = "";
        private string pastaBaseSelecionada = @"C:\LogForge";
        private NumericUpDown nudLimiteEventos;
        private TextBox txtLogPersonalizado;
        private Button btnAddLog;
        private CheckBox chkInfo, chkAviso, chkErro, chkCritico;
        private Button btn24h, btn7dias, btn30dias, btnMesAtual;
        private CheckBox chkUsarFiltroData;
        private List<LogEvento> eventosExtraidos = new List<LogEvento>();

        public FormPrincipal()
        {
            if (!IsAdministrator())
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo(Application.ExecutablePath)
                    {
                        Verb = "runas",
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                    Application.Exit();
                    return;
                }
                catch { }
            }

            this.Text = "LogForge - Modo Depuração";
            this.Size = new Size(1150, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = Color.WhiteSmoke;
            ConstruirInterface();
        }

        private bool IsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void ConstruirInterface()
        {
            TabControl tc = new TabControl() { Dock = DockStyle.Fill };
            TabPage tpConfig = new TabPage("Configuração");
            TabPage tpPreview = new TabPage("Pré-visualização");
            tc.TabPages.Add(tpConfig);
            tc.TabPages.Add(tpPreview);
            this.Controls.Add(tc);

            FlowLayoutPanel flpConfig = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(10), AutoScroll = true };
            tpConfig.Controls.Add(flpConfig);

            // Pasta base
            Panel pnlPasta = new Panel() { Width = 950, Height = 35 };
            Label lblPasta = new Label() { Text = "Pasta base:", Left = 0, Top = 5, Width = 80 };
            TextBox txtPasta = new TextBox() { Left = 85, Top = 3, Width = 350, Text = pastaBaseSelecionada };
            btnSelecionarPasta = new Button() { Text = "Selecionar", Left = 440, Top = 2, Width = 100 };
            btnSelecionarPasta.Click += (s, e) =>
            {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog() { SelectedPath = pastaBaseSelecionada })
                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        pastaBaseSelecionada = fbd.SelectedPath;
                        txtPasta.Text = pastaBaseSelecionada;
                    }
            };
            pnlPasta.Controls.Add(lblPasta);
            pnlPasta.Controls.Add(txtPasta);
            pnlPasta.Controls.Add(btnSelecionarPasta);
            flpConfig.Controls.Add(pnlPasta);

            // Período
            Panel pnlPeriodo = new Panel() { Width = 950, Height = 35 };
            Label lblInicio = new Label() { Text = "Data início:", Left = 0, Top = 5, Width = 70 };
            dtpInicio = new DateTimePicker() { Left = 75, Top = 2, Width = 140, Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm" };
            dtpInicio.Value = DateTime.Now.AddDays(-7);
            Label lblFim = new Label() { Text = "Data fim:", Left = 225, Top = 5, Width = 65 };
            dtpFim = new DateTimePicker() { Left = 295, Top = 2, Width = 140, Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm" };
            dtpFim.Value = DateTime.Now;
            dtpFim.MaxDate = DateTime.Now;
            pnlPeriodo.Controls.Add(lblInicio);
            pnlPeriodo.Controls.Add(dtpInicio);
            pnlPeriodo.Controls.Add(lblFim);
            pnlPeriodo.Controls.Add(dtpFim);
            flpConfig.Controls.Add(pnlPeriodo);

            // Botões rápidos
            Panel pnlRapido = new Panel() { Width = 950, Height = 35 };
            btn24h = new Button() { Text = "Últimas 24h", Left = 0, Width = 100 };
            btn24h.Click += (s, e) => { dtpInicio.Value = DateTime.Now.AddDays(-1); dtpFim.Value = DateTime.Now; };
            btn7dias = new Button() { Text = "Últimos 7 dias", Left = 110, Width = 100 };
            btn7dias.Click += (s, e) => { dtpInicio.Value = DateTime.Now.AddDays(-7); dtpFim.Value = DateTime.Now; };
            pnlRapido.Controls.Add(btn24h);
            pnlRapido.Controls.Add(btn7dias);
            flpConfig.Controls.Add(pnlRapido);

            // Diagnóstico
            chkUsarFiltroData = new CheckBox() { Text = "Usar filtro de data (desmarque para pegar todos os eventos)", Left = 0, Width = 350, Checked = true };
            flpConfig.Controls.Add(chkUsarFiltroData);

            // Filtro de nível
            Panel pnlNivel = new Panel() { Width = 950, Height = 30 };
            chkInfo = new CheckBox() { Text = "Informação", Left = 0, Checked = true };
            chkAviso = new CheckBox() { Text = "Aviso", Left = 100, Checked = true };
            chkErro = new CheckBox() { Text = "Erro", Left = 180, Checked = true };
            chkCritico = new CheckBox() { Text = "Crítico", Left = 260, Checked = true };
            pnlNivel.Controls.Add(chkInfo);
            pnlNivel.Controls.Add(chkAviso);
            pnlNivel.Controls.Add(chkErro);
            pnlNivel.Controls.Add(chkCritico);
            flpConfig.Controls.Add(pnlNivel);

            // Logs
            GroupBox gbLogs = new GroupBox() { Text = "Selecione os logs", Width = 950, Height = 120 };
            chklstLogs = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true };
            foreach (string log in logsDisponiveis) chklstLogs.Items.Add(log, true);
            gbLogs.Controls.Add(chklstLogs);
            flpConfig.Controls.Add(gbLogs);

            // Formato e limite
            Panel pnlFormato = new Panel() { Width = 950, Height = 35 };
            rbtnTXT = new RadioButton() { Text = "TXT", Left = 0, Width = 50, Checked = true };
            rbtnCSV = new RadioButton() { Text = "CSV", Left = 60, Width = 50 };
            Label lblLimite = new Label() { Text = "Máx. eventos por log:", Left = 130, Top = 5, Width = 130 };
            nudLimiteEventos = new NumericUpDown() { Left = 265, Top = 2, Width = 80, Minimum = 100, Maximum = 100000, Value = 10000 };
            pnlFormato.Controls.Add(rbtnTXT);
            pnlFormato.Controls.Add(rbtnCSV);
            pnlFormato.Controls.Add(lblLimite);
            pnlFormato.Controls.Add(nudLimiteEventos);
            flpConfig.Controls.Add(pnlFormato);

            // Botões
            Panel pnlBotoes = new Panel() { Width = 950, Height = 40 };
            btnGerar = new Button() { Text = "Gerar Relatórios (Modo Depuração)", Left = 0, Width = 200, Height = 30, BackColor = Color.SteelBlue, ForeColor = Color.White };
            btnGerar.Click += BtnGerar_Click;
            btnAbrirPasta = new Button() { Text = "Abrir Última Pasta", Left = 210, Width = 130, Height = 30, Enabled = false };
            btnAbrirPasta.Click += BtnAbrirPasta_Click;
            pnlBotoes.Controls.Add(btnGerar);
            pnlBotoes.Controls.Add(btnAbrirPasta);
            flpConfig.Controls.Add(pnlBotoes);

            progressBar = new ProgressBar() { Width = 950, Height = 20, Visible = false };
            flpConfig.Controls.Add(progressBar);

            txtLogStatus = new TextBox() { Width = 950, Height = 300, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, BackColor = Color.Black, ForeColor = Color.LightGreen, Font = new Font("Consolas", 9) };
            flpConfig.Controls.Add(txtLogStatus);

            dgvPreview = new DataGridView() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells };
            tpPreview.Controls.Add(dgvPreview);
        }

        private void BtnGerar_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(pastaBaseSelecionada))
                Directory.CreateDirectory(pastaBaseSelecionada);
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            pastaDestino = Path.Combine(pastaBaseSelecionada, $"Relatorio_{timestamp}");
            Directory.CreateDirectory(pastaDestino);

            txtLogStatus.Clear();
            txtLogStatus.AppendText($"=== MODO DEPURAÇÃO ===\r\n");
            txtLogStatus.AppendText($"Administrador: {IsAdministrator()}\r\n");
            txtLogStatus.AppendText($"Período: {dtpInicio.Value} até {dtpFim.Value}\r\n");
            txtLogStatus.AppendText($"Filtro de data: {chkUsarFiltroData.Checked}\r\n");
            txtLogStatus.AppendText($"Pasta destino: {pastaDestino}\r\n\r\n");

            btnGerar.Enabled = false;
            progressBar.Visible = true;
            progressBar.Style = ProgressBarStyle.Marquee;
            eventosExtraidos.Clear();

            var logsSelecionados = chklstLogs.CheckedItems.Cast<string>().ToList();
            if (logsSelecionados.Count == 0)
            {
                MessageBox.Show("Selecione pelo menos um log.");
                btnGerar.Enabled = true;
                progressBar.Visible = false;
                return;
            }

            System.ComponentModel.BackgroundWorker bw = new System.ComponentModel.BackgroundWorker();
            bw.DoWork += (obj, args) =>
            {
                foreach (string log in logsSelecionados)
                    ProcessarLogDepuracao(log);
            };
            bw.RunWorkerCompleted += (obj, args) =>
            {
                progressBar.Visible = false;
                btnGerar.Enabled = true;
                btnAbrirPasta.Enabled = true;
                txtLogStatus.AppendText("\r\n✅ Processamento concluído!");
                CarregarPreVisualizacao();
            };
            bw.RunWorkerAsync();
        }

        private void ProcessarLogDepuracao(string nomeLog)
        {
            // Teste simples
            string comandoSimples = $"qe \"{nomeLog}\" /c:5 /f:text";
            string saidaSimples = ExecutarWevtutil(comandoSimples);
            this.Invoke(new Action(() => txtLogStatus.AppendText($"\r\n--- TESTE para {nomeLog} (sem filtro) ---\r\n{saidaSimples}\r\n--- FIM TESTE ---\r\n")));

            if (string.IsNullOrWhiteSpace(saidaSimples))
            {
                this.Invoke(new Action(() => txtLogStatus.AppendText($"❌ {nomeLog}: wevtutil não retornou dados! Verifique permissão ou comando.\r\n")));
                return;
            }

            // Comando real
            string query = "*";
            if (chkUsarFiltroData.Checked)
            {
                string dataInicio = dtpInicio.Value.ToString("yyyy-MM-ddTHH:mm:ss");
                string dataFim = dtpFim.Value.ToString("yyyy-MM-ddTHH:mm:ss");
                query = $"*[System[TimeCreated[@SystemTime >= '{dataInicio}' and @SystemTime <= '{dataFim}']]]";
            }
            string comandoReal = $"qe \"{nomeLog}\" /q:\"{query}\" /f:text /rd:true /c:{nudLimiteEventos.Value}";
            string saidaReal = ExecutarWevtutil(comandoReal);

            // Parse e geração do relatório formatado
            var eventos = ParseEventosDetalhado(saidaReal, nomeLog);
            lock (eventosExtraidos) { eventosExtraidos.AddRange(eventos); }

            if (eventos.Count > 0)
                GerarRelatorioFormatado(nomeLog, eventos);
            else
            {
                string arquivo = Path.Combine(pastaDestino, $"{nomeLog}.txt");
                File.WriteAllText(arquivo, "Nenhum evento encontrado com os filtros atuais.", Encoding.UTF8);
            }

            this.Invoke(new Action(() => txtLogStatus.AppendText($"✓ {nomeLog}: {eventos.Count} eventos extraídos e salvo em formato {(rbtnTXT.Checked ? "TXT" : "CSV")}\r\n")));
        }

        private string ExecutarWevtutil(string argumentos)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("wevtutil", argumentos)
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };
                using (Process p = Process.Start(psi))
                {
                    string saida = p.StandardOutput.ReadToEnd();
                    string erro = p.StandardError.ReadToEnd();
                    p.WaitForExit();
                    if (!string.IsNullOrEmpty(erro))
                        return $"ERRO: {erro}\n{saida}";
                    return saida;
                }
            }
            catch (Exception ex)
            {
                return $"EXCEÇÃO: {ex.Message}";
            }
        }

        // ================== NOVOS MÉTODOS DE FORMATAÇÃO ==================

        private List<LogEvento> ParseEventosDetalhado(string raw, string nomeLog)
        {
            List<LogEvento> lista = new List<LogEvento>();
            string[] blocos = raw.Split(new[] { "Event[" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string bloco in blocos)
            {
                LogEvento ev = new LogEvento();
                ev.LogName = nomeLog;
                string dataRaw = ExtrairCampo(bloco, "Date");
                ev.DataHoraFormatada = ConverterDataHora(dataRaw);
                ev.Fonte = ExtrairCampo(bloco, "Source");
                ev.EventoID = ExtrairCampo(bloco, "Event ID");
                string nivelRaw = ExtrairCampo(bloco, "Level");
                ev.Nivel = nivelRaw;
                ev.Usuario = ExtrairCampo(bloco, "User Name");
                if (string.IsNullOrEmpty(ev.Usuario))
                    ev.Usuario = ExtrairCampo(bloco, "User");
                ev.Computador = ExtrairCampo(bloco, "Computer");
                ev.Mensagem = ExtrairDescricao(bloco).Replace("\r\n", " ").Replace("\n", " ");

                // Filtro de nível
                bool incluir = false;
                string nivelLower = nivelRaw.ToLower();
                if (chkInfo.Checked && (nivelLower.Contains("information") || nivelLower.Contains("informação"))) incluir = true;
                else if (chkAviso.Checked && nivelLower.Contains("warning")) incluir = true;
                else if (chkErro.Checked && nivelLower.Contains("error")) incluir = true;
                else if (chkCritico.Checked && nivelLower.Contains("critical")) incluir = true;

                if (incluir && !string.IsNullOrEmpty(ev.DataHoraFormatada))
                    lista.Add(ev);
            }
            return lista;
        }

        private void GerarRelatorioFormatado(string nomeLog, List<LogEvento> eventos)
        {
            string caminho = Path.Combine(pastaDestino, $"{nomeLog}.{(rbtnTXT.Checked ? "txt" : "csv")}");
            using (StreamWriter sw = new StreamWriter(caminho, false, Encoding.UTF8))
            {
                if (rbtnTXT.Checked)
                {
                    // Cabeçalho com larguras fixas
                    string cabecalho = string.Format("{0,-20} {1,-15} {2,-35} {3,-8} {4,-12} {5,-25} {6,-20} {7}",
                        "Data/Hora", "Log", "Fonte", "ID", "Nível", "Usuário", "Computador", "Descrição");
                    sw.WriteLine(cabecalho);
                    sw.WriteLine(new string('-', 180));

                    foreach (var ev in eventos)
                    {
                        string linha = string.Format("{0,-20} {1,-15} {2,-35} {3,-8} {4,-12} {5,-25} {6,-20} {7}",
                            ev.DataHoraFormatada,
                            ev.LogName,
                            Truncar(ev.Fonte, 35),
                            ev.EventoID,
                            TraduzirNivel(ev.Nivel),
                            Truncar(ev.Usuario, 25),
                            Truncar(ev.Computador, 20),
                            Truncar(ev.Mensagem, 80));
                        sw.WriteLine(linha);
                    }
                }
                else // CSV
                {
                    sw.WriteLine("\"Data/Hora\";\"Log\";\"Fonte\";\"ID\";\"Nível\";\"Usuário\";\"Computador\";\"Descrição\"");
                    foreach (var ev in eventos)
                    {
                        sw.WriteLine($"\"{ev.DataHoraFormatada}\";\"{ev.LogName}\";\"{ev.Fonte.Replace("\"", "\"\"")}\";\"{ev.EventoID}\";\"{TraduzirNivel(ev.Nivel)}\";\"{ev.Usuario.Replace("\"", "\"\"")}\";\"{ev.Computador.Replace("\"", "\"\"")}\";\"{ev.Mensagem.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\n", " ")}\"");
                    }
                }
            }
        }

        private string ExtrairCampo(string bloco, string campo)
        {
            var match = Regex.Match(bloco, $@"{campo}:\s*(?<valor>[^\r\n]+)");
            return match.Success ? match.Groups["valor"].Value.Trim() : "";
        }

        private string ExtrairDescricao(string bloco)
        {
            var match = Regex.Match(bloco, @"Description:\s*(?<desc>.*?)(?=\r\n\s*[A-Za-z ]+:|\r\n\r\n|\Z)", RegexOptions.Singleline);
            return match.Success ? match.Groups["desc"].Value.Trim() : "";
        }

        private string ConverterDataHora(string iso)
        {
            if (DateTime.TryParse(iso, out DateTime dt))
                return dt.ToString("dd/MM/yyyy HH:mm:ss");
            return iso;
        }

        private string TraduzirNivel(string nivel)
        {
            if (nivel.Contains("Information")) return "Informação";
            if (nivel.Contains("Warning")) return "Aviso";
            if (nivel.Contains("Error")) return "Erro";
            if (nivel.Contains("Critical")) return "Crítico";
            return nivel;
        }

        private string Truncar(string texto, int max)
        {
            if (string.IsNullOrEmpty(texto)) return "";
            if (texto.Length > max) return texto.Substring(0, max - 3) + "...";
            return texto;
        }

        private void CarregarPreVisualizacao()
        {
            dgvPreview.DataSource = null;
            if (eventosExtraidos.Count > 0)
            {
                var lista = eventosExtraidos.Select(e => new
                {
                    Data = e.DataHoraFormatada,
                    Nível = TraduzirNivel(e.Nivel),
                    Fonte = e.Fonte,
                    ID = e.EventoID,
                    Usuário = e.Usuario,
                    Computador = e.Computador,
                    Mensagem = e.Mensagem.Length > 100 ? e.Mensagem.Substring(0, 100) + "..." : e.Mensagem
                }).Take(500).ToList();
                dgvPreview.DataSource = lista;
            }
        }

        private void BtnAbrirPasta_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(pastaDestino) && Directory.Exists(pastaDestino))
                Process.Start("explorer.exe", pastaDestino);
            else
                MessageBox.Show("Pasta não encontrada.");
        }

        // Classe aninhada para armazenar os dados do evento
        private class LogEvento
        {
            public string LogName { get; set; }
            public string DataHoraFormatada { get; set; }
            public string Nivel { get; set; }
            public string Fonte { get; set; }
            public string EventoID { get; set; }
            public string Mensagem { get; set; }
            public string Usuario { get; set; }
            public string Computador { get; set; }
        }
    }
}