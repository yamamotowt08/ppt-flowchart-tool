using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace FlowchartTool
{
    static class Program
    {
        internal const string Title = "フローチャート作成ツール";
        // 作業用のフォルダは各自のパソコンに作る（共有フォルダに書き込まないように）
        internal static readonly string LocalDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FlowchartTool");
        internal static string LibDir;

        [STAThread]
        static void Main()
        {
            try { LibDir = ExtractLibraries(); }
            catch (Exception ex)
            {
                MessageBox.Show("起動の準備ができませんでした。\n\n" + ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // EXE の横に WebView2 の DLL がないときは、取り出した DLL を使う
            AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
            {
                string path = Path.Combine(LibDir, new AssemblyName(e.Name).Name + ".dll");
                return File.Exists(path) ? Assembly.LoadFrom(path) : null;
            };
            Run();
        }

        // WebView2 の型を使うのは、DLL の読み込み先を決めたあと（このメソッドの中）から
        [MethodImpl(MethodImplOptions.NoInlining)]
        static void Run()
        {
            CoreWebView2Environment.SetLoaderDllFolderPath(LibDir);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        // EXE の中に入れてある WebView2 の DLL を、作業用のフォルダに取り出す（版と CPU の種類ごとのフォルダ）
        static string ExtractLibraries()
        {
            string arch = RuntimeInformation.ProcessArchitecture == Architecture.Arm64 ? "arm64"
                : Environment.Is64BitProcess ? "x64" : "x86";
            string dir = Path.Combine(LocalDir, "lib", Assembly.GetExecutingAssembly().GetName().Version + "-" + arch);
            Directory.CreateDirectory(dir);
            ExtractResource("lib.Microsoft.Web.WebView2.Core.dll", Path.Combine(dir, "Microsoft.Web.WebView2.Core.dll"));
            ExtractResource("lib.Microsoft.Web.WebView2.WinForms.dll", Path.Combine(dir, "Microsoft.Web.WebView2.WinForms.dll"));
            ExtractResource("native." + arch + ".WebView2Loader.dll", Path.Combine(dir, "WebView2Loader.dll"));
            return dir;
        }

        internal static void ExtractResource(string name, string path)
        {
            using (Stream src = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (src == null) throw new FileNotFoundException("アプリの中に " + name + " が見つかりません。");
                if (File.Exists(path) && new FileInfo(path).Length == src.Length) return; // 取り出し済み
                string tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                using (Stream dst = File.Create(tmp)) src.CopyTo(dst);
                try
                {
                    if (File.Exists(path)) File.Delete(path);
                    File.Move(tmp, path);
                }
                catch (IOException) when (File.Exists(path))
                {
                    File.Delete(tmp); // ほかに起動中のアプリが使っているときは、そのファイルをそのまま使う
                }
            }
        }
    }

    // フローチャート作成ツール（HTML）を、Edge の部品（WebView2）で表示するアプリ。
    // 「PowerPoint で保存」などで作ったファイルは、Windows の保存の画面で保存先を選んで保存する。
    sealed class MainForm : Form
    {
        const string Host = "flowchart-tool.local"; // アプリの中だけで使うアドレス（書いた内容の保存先もこれで決まる）
        const string ToolName = "flowchart.html";
        const string RuntimeUrl = "https://developer.microsoft.com/microsoft-edge/webview2/";

        readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
        string lastSaveDir;

        public MainForm()
        {
            Text = Program.Title;
            Size = new Size(1400, 900);
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { /* アイコンがなくても動く */ }
            Controls.Add(web);
            Load += async (s, e) => await StartAsync();
        }

        async Task StartAsync()
        {
            try
            {
                string appDir = Path.Combine(Program.LocalDir, "app");
                Directory.CreateDirectory(appDir);
                // ツール本体は EXE の中に入っている。毎回取り出し直して、EXE と同じ版を使う
                string target = Path.Combine(appDir, ToolName);
                if (File.Exists(target)) File.Delete(target);
                Program.ExtractResource("tool.html", target);

                var env = await CoreWebView2Environment.CreateAsync(null, Path.Combine(Program.LocalDir, "WebView2"));
                await web.EnsureCoreWebView2Async(env);
                CoreWebView2 core = web.CoreWebView2;
                core.SetVirtualHostNameToFolderMapping(Host, appDir, CoreWebView2HostResourceAccessKind.Allow);

                // PowerPoint・画像の保存は、Windows の保存の画面で場所と名前を選ぶ
                core.DownloadStarting += OnDownloadStarting;
                // ツール以外のページ（リンク）は、ふだんのブラウザで開く
                core.NewWindowRequested += (s, e) => { e.Handled = true; OpenInBrowser(e.Uri); };
                core.NavigationStarting += (s, e) =>
                {
                    if (!IsTool(e.Uri)) { e.Cancel = true; OpenInBrowser(e.Uri); }
                };
                core.DocumentTitleChanged += (s, e) => Text = Program.Title;

                core.Navigate("https://" + Host + "/" + ToolName);
            }
            catch (WebView2RuntimeNotFoundException)
            {
                if (MessageBox.Show(this,
                        "このパソコンには、アプリが使う「Microsoft Edge WebView2 ランタイム」が入っていません。\n\n" +
                        "ダウンロードのページを開きますか？（入れるには、パソコン管理担当の許可が必要なことがあります）",
                        Program.Title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    OpenInBrowser(RuntimeUrl);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "起動できませんでした。\n\n" + ex.Message, Program.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        void OnDownloadStarting(object sender, CoreWebView2DownloadStartingEventArgs e)
        {
            CoreWebView2Deferral deferral = e.GetDeferral();
            string suggested = Path.GetFileName(e.ResultFilePath);
            // 保存の画面は、イベントの外で出す
            BeginInvoke((Action)(() =>
            {
                try
                {
                    using (var dlg = new SaveFileDialog
                    {
                        Title = "保存する場所を選ぶ",
                        FileName = suggested,
                        Filter = FilterFor(suggested),
                        OverwritePrompt = true,
                        InitialDirectory = lastSaveDir ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    })
                    {
                        if (dlg.ShowDialog(this) == DialogResult.OK)
                        {
                            e.ResultFilePath = dlg.FileName;
                            e.Handled = true; // 画面右上のダウンロードの表示は出さない
                            lastSaveDir = Path.GetDirectoryName(dlg.FileName);
                        }
                        else
                        {
                            e.Cancel = true;
                        }
                    }
                }
                finally
                {
                    deferral.Complete();
                }
            }));
        }

        static string FilterFor(string name)
        {
            string ext = Path.GetExtension(name ?? "").ToLowerInvariant();
            if (ext == ".pptx") return "PowerPoint プレゼンテーション (*.pptx)|*.pptx|すべてのファイル (*.*)|*.*";
            if (ext == ".png") return "PNG 画像 (*.png)|*.png|すべてのファイル (*.*)|*.*";
            return "すべてのファイル (*.*)|*.*";
        }

        static bool IsTool(string uri) =>
            uri != null && uri.StartsWith("https://" + Host + "/", StringComparison.OrdinalIgnoreCase);

        static void OpenInBrowser(string uri)
        {
            if (uri == null || !(uri.StartsWith("https://") || uri.StartsWith("http://"))) return;
            try { Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true }); } catch { /* 開けなくても続ける */ }
        }
    }
}
