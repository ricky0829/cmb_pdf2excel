// 招商银行交易流水提取工具 - C# WinForms 原生窗口（单文件版）
// 编译: csc /target:winexe /out:cmb_pdf2excel.exe /resource:worker.exe gui.cs
// worker.exe 作为内嵌资源打包进本程序，运行时自动释放到临时目录后调用

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;

public class MainForm : Form
{
    private TextBox txtPdf, txtXlsx, txtLog;
    private Button btnBrowsePdf, btnBrowseXlsx, btnRun, btnOpen, btnClear;
    private Label lblStatus;
    private Process worker;

    public MainForm()
    {
        Text = "招商银行交易流水提取工具";
        Size = new Size(700, 500);
        MinimumSize = new Size(580, 420);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft YaHei UI", 9f);
        BuildUI();
    }

    private void BuildUI()
    {
        int pad = 12, y = pad;
        int cw = ClientSize.Width;   // 当前客户区宽度
        int btnW = 80, btnH = 32;    // 浏览按钮与下方按钮同尺寸
        int txtRight = cw - pad - btnW - 8;  // 文本框右边界

        // ---- PDF 文件 ----
        var lblPdf = new Label { Text = "PDF 文件", Location = new Point(pad, y + 8), AutoSize = true };
        txtPdf = new TextBox
        {
            Location = new Point(pad + 70, y + 5), Width = txtRight - (pad + 70),
            ReadOnly = true, BackColor = SystemColors.Window,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        btnBrowsePdf = new Button
        {
            Text = "浏览…", Location = new Point(cw - pad - btnW, y), Width = btnW, Height = btnH,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnBrowsePdf.Click += (s, e) => BrowsePdf();
        Controls.AddRange(new Control[] { lblPdf, txtPdf, btnBrowsePdf });
        y += 40;

        // ---- 输出 Excel ----
        var lblXlsx = new Label { Text = "输出 Excel", Location = new Point(pad, y + 8), AutoSize = true };
        txtXlsx = new TextBox
        {
            Location = new Point(pad + 70, y + 5), Width = txtRight - (pad + 70),
            ReadOnly = true, BackColor = SystemColors.Window,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        btnBrowseXlsx = new Button
        {
            Text = "浏览…", Location = new Point(cw - pad - btnW, y), Width = btnW, Height = btnH,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        btnBrowseXlsx.Click += (s, e) => BrowseXlsx();
        Controls.AddRange(new Control[] { lblXlsx, txtXlsx, btnBrowseXlsx });
        y += 42;

        // ---- 按钮行 ----
        btnRun = new Button { Text = "开始提取", Location = new Point(pad, y), Width = 90, Height = 32 };
        btnRun.Click += (s, e) => Run();
        btnOpen = new Button { Text = "打开输出目录", Location = new Point(pad + 100, y), Width = 100, Height = 32, Enabled = false };
        btnOpen.Click += (s, e) => OpenDir();
        btnClear = new Button { Text = "清空日志", Location = new Point(pad + 210, y), Width = 80, Height = 32 };
        btnClear.Click += (s, e) => { txtLog.Clear(); };
        lblStatus = new Label { Text = "就绪 — 请选择 PDF 文件", Location = new Point(pad + 310, y + 8), AutoSize = true, ForeColor = Color.Gray };
        Controls.AddRange(new Control[] { btnRun, btnOpen, btnClear, lblStatus });
        y += 42;

        // ---- 日志 ----
        var lblLog = new Label { Text = "运行日志", Location = new Point(pad, y), AutoSize = true };
        y += 20;
        txtLog = new TextBox
        {
            Location = new Point(pad, y),
            Size = new Size(cw - pad * 2, ClientSize.Height - y - pad),
            Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
            BackColor = Color.FromArgb(30, 30, 46), ForeColor = Color.FromArgb(166, 227, 161),
            Font = new Font("Consolas", 9f), WordWrap = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        Controls.AddRange(new Control[] { lblLog, txtLog });
    }

    // ---------- 文件选择 ----------

    private void BrowsePdf()
    {
        using (var dlg = new OpenFileDialog())
        {
            dlg.Title = "选择招商银行交易流水 PDF";
            dlg.Filter = "PDF 文件 (*.pdf)|*.pdf|所有文件 (*.*)|*.*";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtPdf.Text = dlg.FileName;
                txtXlsx.Text = Path.ChangeExtension(dlg.FileName, ".xlsx");
            }
        }
    }

    private void BrowseXlsx()
    {
        using (var dlg = new SaveFileDialog())
        {
            dlg.Title = "保存 Excel 文件";
            dlg.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
            dlg.FileName = Path.GetFileNameWithoutExtension(txtPdf.Text) + ".xlsx";
            if (dlg.ShowDialog() == DialogResult.OK)
                txtXlsx.Text = dlg.FileName;
        }
    }

    // ---------- 释放内嵌 worker.exe ----------

    // 从本程序内嵌资源中释放 worker.exe 到临时目录。
    // 以内容哈希命名子目录，保证不同版本互不覆盖、且已存在时跳过写入（避免文件占用）。
    // 若资源未内嵌（开发调试场景）或释放失败，返回 null 由调用方回退到同目录查找。
    private string EnsureWorker()
    {
        try
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream rs = asm.GetManifestResourceStream("worker.exe"))
            {
                if (rs == null) return null;
                byte[] data;
                using (var ms = new MemoryStream())
                { rs.CopyTo(ms); data = ms.ToArray(); }

                string hash;
                using (var sha = SHA1.Create())
                {
                    byte[] h = sha.ComputeHash(data);
                    hash = BitConverter.ToString(h).Replace("-", "").Substring(0, 10);
                }

                string dir = Path.Combine(Path.GetTempPath(), "cmb_pdf2excel", hash);
                Directory.CreateDirectory(dir);
                string workerPath = Path.Combine(dir, "worker.exe");
                if (!File.Exists(workerPath))
                    File.WriteAllBytes(workerPath, data);
                return workerPath;
            }
        }
        catch
        {
            return null;
        }
    }

    // ---------- 提取流程 ----------

    private void Run()
    {
        string pdf = txtPdf.Text.Trim();
        string xlsx = txtXlsx.Text.Trim();
        if (string.IsNullOrEmpty(pdf) || !File.Exists(pdf))
        { MessageBox.Show("请先选择有效的 PDF 文件。", "提示"); return; }
        if (string.IsNullOrEmpty(xlsx))
        { MessageBox.Show("请指定输出 Excel 路径。", "提示"); return; }

        // 定位 worker.exe：优先从内嵌资源释放，失败则回退到同目录
        string workerPath = EnsureWorker();
        if (workerPath == null || !File.Exists(workerPath))
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            workerPath = Path.Combine(exeDir, "worker.exe");
        }
        if (!File.Exists(workerPath))
        {
            MessageBox.Show("未找到 worker.exe，程序无法运行。", "错误");
            return;
        }

        btnRun.Enabled = false;
        btnOpen.Enabled = false;
        lblStatus.Text = "正在提取，请稍候…";
        lblStatus.ForeColor = Color.DodgerBlue;

        worker = new Process();
        worker.StartInfo.FileName = workerPath;
        worker.StartInfo.Arguments = "\"" + pdf + "\" \"" + xlsx + "\"";
        worker.StartInfo.UseShellExecute = false;
        worker.StartInfo.RedirectStandardOutput = true;
        worker.StartInfo.RedirectStandardError = true;
        worker.StartInfo.CreateNoWindow = true;
        worker.StartInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
        worker.StartInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;
        worker.EnableRaisingEvents = true;
        worker.OutputDataReceived += (s, e) => { if (e.Data != null) SafeLog(e.Data); };
        worker.ErrorDataReceived += (s, e) => { if (e.Data != null) SafeLog(e.Data); };
        worker.Exited += (s, e) => OnWorkerExit(worker.ExitCode, xlsx);
        worker.Start();
        worker.BeginOutputReadLine();
        worker.BeginErrorReadLine();
    }

    private void SafeLog(string line)
    {
        if (InvokeRequired) { Invoke(new Action<string>(SafeLog), line); return; }
        txtLog.AppendText(line + Environment.NewLine);
    }

    private void OnWorkerExit(int code, string xlsx)
    {
        if (InvokeRequired) { Invoke(new Action<int, string>(OnWorkerExit), code, xlsx); return; }
        if (code == 0)
        {
            lblStatus.Text = "提取完成 ✓";
            lblStatus.ForeColor = Color.Green;
            btnOpen.Enabled = true;
            btnOpen.Tag = Path.GetDirectoryName(xlsx);
            MessageBox.Show("已成功导出到:\n" + xlsx, "完成");
        }
        else
        {
            lblStatus.Text = "出错";
            lblStatus.ForeColor = Color.Red;
        }
        btnRun.Enabled = true;
    }

    private void OpenDir()
    {
        string dir = btnOpen.Tag as string;
        if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
            Process.Start("explorer.exe", dir);
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
