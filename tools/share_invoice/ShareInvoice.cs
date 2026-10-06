// ShareInvoice.exe — opens the Windows Share window with one file and a text message.
// The person picks the app (e.g. WhatsApp) and the contact there; nothing is ever sent automatically.
//
// Usage:  ShareInvoice.exe <file> [<utf-8 text file with the message>] [<title>]
// Output: "TARGET:<app name>" when an app was chosen, "CANCELLED" otherwise, "ERROR:<details>" on failure.
//
// Build (no SDK needed, .NET Framework 4 is part of Windows 10/11): see build.cmd next to this file.
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;

[ComImport, Guid("3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IDataTransferManagerInterop
{
    IntPtr GetForWindow([In] IntPtr appWindow, [In] ref Guid riid);
    void ShowShareUIForWindow(IntPtr appWindow);
}

static class ShareInvoice
{
    [DllImport("combase.dll", PreserveSig = false)]
    static extern void RoGetActivationFactory([MarshalAs(UnmanagedType.HString)] string classId, [In] ref Guid iid,
                                              [MarshalAs(UnmanagedType.IUnknown)] out object factory);

    static readonly Guid IID_IDataTransferManager = new Guid("a5caee9b-8708-49d1-8d36-67d25a8da00c");
    const int TimeoutSeconds = 300;

    [STAThread]
    static int Main(string[] args)
    {
        // A GUI-subsystem exe cannot set Console.OutputEncoding; write UTF-8 to the redirected stdout instead.
        var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        if (args.Length < 1)
        {
            stdout.WriteLine("ERROR:usage: ShareInvoice.exe <file> [<message.txt>] [<title>]");
            return 2;
        }
        try
        {
            string path = Path.GetFullPath(args[0]);
            string text = args.Length > 1 && File.Exists(args[1]) ? File.ReadAllText(args[1], Encoding.UTF8) : "";
            string title = args.Length > 2 ? args[2] : Path.GetFileNameWithoutExtension(path);
            string outcome = Run(path, text, title);
            stdout.WriteLine(outcome);
            return outcome.StartsWith("TARGET:") ? 0 : 1;
        }
        catch (Exception e)
        {
            stdout.WriteLine("ERROR:" + e.Message.Replace("\r", " ").Replace("\n", " "));
            return 3;
        }
    }

    // Blocks on a WinRT async operation without System.Runtime.WindowsRuntime's AsTask, which would
    // need the Windows SDK's union metadata to compile.
    static T Wait<T>(Windows.Foundation.IAsyncOperation<T> op)
    {
        var info = (Windows.Foundation.IAsyncInfo)op;
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (info.Status == Windows.Foundation.AsyncStatus.Started)
        {
            if (DateTime.UtcNow > deadline) throw new TimeoutException("Windows did not open the file.");
            System.Threading.Thread.Sleep(10);
        }
        if (info.Status != Windows.Foundation.AsyncStatus.Completed)
            throw info.ErrorCode ?? new InvalidOperationException("Windows could not open the file.");
        return op.GetResults();
    }

    static string Run(string path, string text, string title)
    {
        StorageFile file = Wait(StorageFile.GetFileFromPathAsync(path));
        string outcome = "CANCELLED";

        // The Share window has to belong to a window of ours: a small centred "pick WhatsApp" note.
        var form = new Form
        {
            Text = "Share invoice",
            Width = 460,
            Height = 140,
            StartPosition = FormStartPosition.CenterScreen,
            TopMost = true,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = true,
        };
        form.Controls.Add(new Label
        {
            Text = "Choose WhatsApp in the Share window, then pick the contact and press Send.\n" +
                   Path.GetFileName(path),
            Dock = DockStyle.Fill,
            TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
        });

        var timer = new Timer { Interval = TimeoutSeconds * 1000 };
        timer.Tick += (s, e) => form.Close();

        // Windows only shows the Share window for the foreground window, so come to the front first
        // and open Share once the window is really active.
        Action openShare = () =>
        {
            Log("opening share; foreground=" + (GetForegroundWindow() == form.Handle));
            Guid interopIid = typeof(IDataTransferManagerInterop).GUID;
            object factory;
            RoGetActivationFactory("Windows.ApplicationModel.DataTransfer.DataTransferManager", ref interopIid, out factory);
            var interop = (IDataTransferManagerInterop)factory;

            Guid dtmIid = IID_IDataTransferManager;
            IntPtr p = interop.GetForWindow(form.Handle, ref dtmIid);
            var dtm = (DataTransferManager)Marshal.GetObjectForIUnknown(p);
            Marshal.Release(p);

            dtm.DataRequested += (sender, a) =>
            {
                Log("data requested");
                DataPackage data = a.Request.Data;
                data.Properties.Title = title;
                if (!string.IsNullOrEmpty(text)) data.SetText(text);
                data.SetStorageItems(new List<IStorageItem> { file });
            };
            dtm.TargetApplicationChosen += (sender, a) =>
            {
                outcome = "TARGET:" + a.ApplicationName;
                // Keep the process alive briefly: the target app reads the file/text from us.
                form.BeginInvoke((Action)(() =>
                {
                    var closer = new Timer { Interval = 15000 };
                    closer.Tick += (s2, e2) => form.Close();
                    closer.Start();
                    form.WindowState = FormWindowState.Minimized;
                }));
            };
            interop.ShowShareUIForWindow(form.Handle);
            Log("share window requested");
            timer.Start();
        };
        form.Shown += (s, e) =>
        {
            ComeToFront(form.Handle);
            var delay = new Timer { Interval = 200 };
            delay.Tick += (s2, e2) => { delay.Stop(); openShare(); };
            delay.Start();
        };
        Application.Run(form);
        return outcome;
    }

    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);

    // Windows lets a window take the foreground only in some situations; when it refuses, share the
    // foreground thread's input state for a moment (and tap Alt, which lifts the foreground lock).
    static void ComeToFront(IntPtr hwnd)
    {
        if (SetForegroundWindow(hwnd) && GetForegroundWindow() == hwnd) return;
        uint fgThread = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero);
        uint ourThread = GetCurrentThreadId();
        bool attached = fgThread != 0 && fgThread != ourThread && AttachThreadInput(ourThread, fgThread, true);
        try
        {
            const byte VK_MENU = 0x12; const uint KEYUP = 0x0002;
            keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
            keybd_event(VK_MENU, 0, KEYUP, UIntPtr.Zero);
            BringWindowToTop(hwnd);
            SetForegroundWindow(hwnd);
        }
        finally
        {
            if (attached) AttachThreadInput(ourThread, fgThread, false);
        }
    }

    // Optional diagnostics: set SHARE_INVOICE_LOG to a file path.
    static void Log(string message)
    {
        string path = Environment.GetEnvironmentVariable("SHARE_INVOICE_LOG");
        if (string.IsNullOrEmpty(path)) return;
        try { File.AppendAllText(path, DateTime.Now.ToString("HH:mm:ss.fff ") + message + Environment.NewLine); }
        catch (IOException) { }
    }
}
