using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace tool_type_instead_of_paste;

public partial class MainWindow : Window
{
    #region Win32 SendInput

    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_UNICODE = 0x0004;
    private const ushort VK_RETURN = 0x0D;

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint Type;
        public InputUnion Data;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    #endregion

    private readonly ObservableCollection<string> _history = new();
    private DispatcherTimer? _timer;
    private int _countdown;

    public MainWindow()
    {
        InitializeComponent();
        HistoryListBox.ItemsSource = _history;
    }

    private void SendButton_Click(object sender, RoutedEventArgs e)
    {
        var text = InputTextBox.Text;
        if (string.IsNullOrEmpty(text)) return;

        _timer?.Stop();
        SendButton.IsEnabled = false;
        _countdown = 3;
        SendButton.Content = _countdown.ToString();

        _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _timer.Tick += (s, args) =>
        {
            _countdown--;
            if (_countdown > 0)
            {
                SendButton.Content = _countdown.ToString();
            }
            else
            {
                _timer.Stop();
                SendButton.Content = "3秒後に送出";
                SendButton.IsEnabled = true;

                if (!_history.Contains(text))
                    _history.Insert(0, text);

                TypeText(text);
            }
        };
        _timer.Start();
    }

    private void HistoryListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (HistoryListBox.SelectedItem is string selected)
        {
            InputTextBox.Text = selected;
            HistoryListBox.SelectedItem = null;
        }
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        _history.Clear();
    }

    private static void TypeText(string text)
    {
        var inputs = new List<INPUT>();

        foreach (var c in text)
        {
            if (c == '\r') continue;

            if (c == '\n')
            {
                inputs.Add(CreateVkInput(VK_RETURN, 0));
                inputs.Add(CreateVkInput(VK_RETURN, KEYEVENTF_KEYUP));
            }
            else
            {
                inputs.Add(CreateUnicodeInput(c, KEYEVENTF_UNICODE));
                inputs.Add(CreateUnicodeInput(c, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP));
            }
        }

        var array = inputs.ToArray();
        if (array.Length == 0) return;

        var sent = SendInput((uint)array.Length, array, Marshal.SizeOf<INPUT>());
        if (sent < array.Length)
            MessageBox.Show(
                "キーボード入力の送出に失敗しました。\n対象ウィンドウが高い権限で実行されている可能性があります。",
                "送出エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private static INPUT CreateUnicodeInput(char c, uint flags) => new()
    {
        Type = INPUT_KEYBOARD,
        Data = new InputUnion
        {
            ki = new KEYBDINPUT
            {
                wVk = 0,
                wScan = (ushort)c,
                dwFlags = flags,
                time = 0,
                dwExtraInfo = IntPtr.Zero
            }
        }
    };

    private static INPUT CreateVkInput(ushort vk, uint flags) => new()
    {
        Type = INPUT_KEYBOARD,
        Data = new InputUnion
        {
            ki = new KEYBDINPUT
            {
                wVk = vk,
                wScan = 0,
                dwFlags = flags,
                time = 0,
                dwExtraInfo = IntPtr.Zero
            }
        }
    };
}
