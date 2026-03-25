using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Input;
using System.Windows.Media.Animation;

namespace MyFirstWPFApp
{
    public partial class MainWindow : Window
    {
        private const int AutoClickerHotkeyId = 0x4910;
        private const int ScriptToggleHotkeyId = 0x4911;
        private const int ScriptPauseHotkeyId = 0x4912;
        private const int WmHotkey = 0x0312;

        private const uint ModAlt = 0x0001;
        private const uint ModControl = 0x0002;
        private const uint ModShift = 0x0004;
        private const uint ModWin = 0x0008;

        private const int WhKeyboardLl = 13;
        private const int WmKeydown = 0x0100;
        private const int WmKeyup = 0x0101;
        private const int WmSyskeydown = 0x0104;
        private const int WmSyskeyup = 0x0105;

        private const uint InputMouse = 0;
        private const uint InputKeyboard = 1;
        private const uint MouseeventfLeftdown = 0x0002;
        private const uint MouseeventfLeftup = 0x0004;
        private const uint MouseeventfRightdown = 0x0008;
        private const uint MouseeventfRightup = 0x0010;
        private const uint KeyeventfKeyup = 0x0002;
        private const uint KeyeventfUnicode = 0x0004;

        private HwndSource? _source;
        private bool _autoClickerHotkeyRegistered;
        private bool _scriptToggleHotkeyRegistered;
        private bool _scriptPauseHotkeyRegistered;

        private CancellationTokenSource? _autoClickerCts;
        private bool _isAutoClickerRunning;

        private IntPtr _keyboardHookHandle = IntPtr.Zero;
        private HookProc? _keyboardProc;
        private bool _isRecording;
        private readonly Stopwatch _recordingStopwatch = new();
        private long _lastRecordedMs;

        private CancellationTokenSource? _scriptPlaybackCts;
        private bool _isScriptPaused;
        private List<AutomationProfile> _profiles = new();
        private const double SimpleModeMinHeight = 430;
        private const double AdvancedModeMinHeight = 640;

        private enum ClickMode
        {
            Left,
            Right,
            DoubleLeft,
        }

        public MainWindow()
        {
            InitializeComponent();
            PopulateHotkeyKeys();
            HotkeyComboBox.SelectedItem = Key.F6;
            CtrlModifierCheckBox.IsChecked = true;
            ScriptToggleHotkeyComboBox.SelectedItem = Key.F7;
            ScriptPauseHotkeyComboBox.SelectedItem = Key.F8;
            ScriptCtrlModifierCheckBox.IsChecked = true;
            ClickTypeComboBox.SelectedIndex = 0;
            UseCurrentCursorCheckBox.IsChecked = true;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _source = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
            _source?.AddHook(WndProc);

            ApplyHotkeyRegistration();
            ApplyScriptHotkeyRegistration();
            StartKeyboardHook();
            RefreshPresets();
            UpdateWindowSizeMode();
        }

        private void Window_Closed(object? sender, EventArgs e)
        {
            StopAutoClicker();
            StopScriptPlayback();
            StopRecording();
            StopKeyboardHook();
            UnregisterAutoClickerHotkeyIfNeeded();
            UnregisterScriptHotkeysIfNeeded();

            if (_source is not null)
            {
                _source.RemoveHook(WndProc);
                _source = null;
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
                MaximizeOrRestore();
            else
                DragMove();
        }

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void Maximize_Click(object sender, RoutedEventArgs e)
        {
            MaximizeOrRestore();
        }

        private void MaximizeOrRestore()
        {
            WindowState = WindowState == WindowState.Maximized
                ? WindowState.Normal
                : WindowState.Maximized;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CaptureMousePosition_Click(object sender, RoutedEventArgs e)
        {
            if (GetCursorPos(out var point))
            {
                ClickXTextBox.Text = point.X.ToString(CultureInfo.InvariantCulture);
                ClickYTextBox.Text = point.Y.ToString(CultureInfo.InvariantCulture);
                MousePositionStatusText.Text = $"Positie gezet naar X={point.X}, Y={point.Y}";
            }
            else
            {
                MousePositionStatusText.Text = "Kon muispositie niet ophalen.";
            }
        }

        private void AdvancedExpander_Expanded(object sender, RoutedEventArgs e)
        {
            ContentScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            AnimateAdvancedPanel(0.0, 1.0, 180);
            UpdateWindowSizeMode();
        }

        private void AdvancedExpander_Collapsed(object sender, RoutedEventArgs e)
        {
            // Hide scrollbar temporarily before collapse animation
            ContentScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Hidden;
            AnimateAdvancedPanel(1.0, 0.0, 140);
            UpdateWindowSizeMode();
        }

        private void AnimateAdvancedPanel(double from, double to, int milliseconds)
        {
            var animation = new DoubleAnimation
            {
                From = from,
                To = to,
                Duration = TimeSpan.FromMilliseconds(milliseconds),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut },
            };

            AdvancedContentPanel.BeginAnimation(OpacityProperty, animation);
        }

        private void UpdateWindowSizeMode()
        {
            bool isAdvanced = AdvancedExpander?.IsExpanded == true;

            if (isAdvanced)
            {
                SizeToContent = SizeToContent.Manual;
                MinHeight = AdvancedModeMinHeight;
                if (Height < AdvancedModeMinHeight)
                {
                    Height = AdvancedModeMinHeight;
                }
            }
            else
            {
                SizeToContent = SizeToContent.Height;
                MinHeight = SimpleModeMinHeight;
            }
        }

        private async void CaptureMousePositionWithDelay_Click(object sender, RoutedEventArgs e)
        {
            MousePositionStatusText.Text = "Plaats je muis op de gewenste plek... 3";
            await Task.Delay(1000);
            MousePositionStatusText.Text = "Plaats je muis op de gewenste plek... 2";
            await Task.Delay(1000);
            MousePositionStatusText.Text = "Plaats je muis op de gewenste plek... 1";
            await Task.Delay(1000);

            if (GetCursorPos(out var point))
            {
                ClickXTextBox.Text = point.X.ToString(CultureInfo.InvariantCulture);
                ClickYTextBox.Text = point.Y.ToString(CultureInfo.InvariantCulture);
                MousePositionStatusText.Text = $"Positie gezet naar X={point.X}, Y={point.Y}";
            }
            else
            {
                MousePositionStatusText.Text = "Kon muispositie niet ophalen.";
            }
        }

        private void ApplyHotkey_Click(object sender, RoutedEventArgs e)
        {
            if (ApplyHotkeyRegistration())
            {
                AutoClickerStatusText.Text = "Status: uit (hotkey actief)";
            }
        }

        private void ApplyScriptHotkeys_Click(object sender, RoutedEventArgs e)
        {
            if (ApplyScriptHotkeyRegistration())
            {
                KeyboardStatusText.Text = "Status: script hotkeys actief";
            }
        }

        private void ToggleAutoClickerButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleAutoClicker();
        }

        private async void PlayScriptButton_Click(object sender, RoutedEventArgs e)
        {
            if (_scriptPlaybackCts is not null)
            {
                KeyboardStatusText.Text = "Status: script draait al";
                return;
            }

            if (!TryParseScript(ScriptTextBox.Text, out var commands, out var parseError))
            {
                KeyboardStatusText.Text = parseError;
                return;
            }

            _scriptPlaybackCts = new CancellationTokenSource();
            _isScriptPaused = false;
            bool loopEnabled = LoopScriptCheckBox.IsChecked == true;
            KeyboardStatusText.Text = loopEnabled ? "Status: script loopen" : "Status: script afspelen";

            try
            {
                do
                {
                    await PlayCommandsAsync(commands, _scriptPlaybackCts.Token);
                }
                while (loopEnabled && _scriptPlaybackCts is not null && !_scriptPlaybackCts.IsCancellationRequested);

                if (_scriptPlaybackCts is not null)
                {
                    KeyboardStatusText.Text = "Status: script klaar";
                }
            }
            catch (OperationCanceledException)
            {
                KeyboardStatusText.Text = "Status: afspelen gestopt";
            }
            finally
            {
                _scriptPlaybackCts?.Dispose();
                _scriptPlaybackCts = null;
            }
        }

        private void StopScriptButton_Click(object sender, RoutedEventArgs e)
        {
            StopScriptPlayback();
            KeyboardStatusText.Text = "Status: afspelen gestopt";
        }

        private void RecordToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isRecording)
            {
                StopRecording();
                KeyboardStatusText.Text = "Status: opname gestopt";
                return;
            }

            _isRecording = true;
            _recordingStopwatch.Restart();
            _lastRecordedMs = 0;
            ScriptTextBox.Clear();
            RecordToggleButton.Content = "Stop opname";
            KeyboardStatusText.Text = "Status: opnemen...";
        }

        private void ToggleAutoClicker()
        {
            if (_isAutoClickerRunning)
            {
                StopAutoClicker();
                return;
            }

            if (!TryReadAutoClickerConfig(out var config))
            {
                return;
            }

            _autoClickerCts = new CancellationTokenSource();
            _isAutoClickerRunning = true;
            ToggleAutoClickerButton.Content = "Stop autoclicker";
            AutoClickerStatusText.Text = "Status: aan";

            _ = RunAutoClickerAsync(config, _autoClickerCts.Token);
        }

        private void StopAutoClicker()
        {
            _autoClickerCts?.Cancel();
            _autoClickerCts?.Dispose();
            _autoClickerCts = null;

            _isAutoClickerRunning = false;
            ToggleAutoClickerButton.Content = "Start autoclicker";
            AutoClickerStatusText.Text = "Status: uit";
        }

        private bool TryReadAutoClickerConfig(out AutoClickerConfig config)
        {
            config = default;

            int x = 500;
            int y = 500;
            int intervalMs;

            if (!double.TryParse(SimpleIntervalSecondsTextBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double secondsPerClick) ||
                secondsPerClick <= 0)
            {
                AutoClickerStatusText.Text = "Status: vul geldige seconden per klik in";
                return false;
            }

            intervalMs = Math.Max(1, (int)Math.Round(secondsPerClick * 1000));

            if (ClickTypeComboBox.SelectedIndex < 0)
            {
                AutoClickerStatusText.Text = "Status: kies een kliktype";
                return false;
            }

            ClickMode clickMode = ClickTypeComboBox.SelectedIndex switch
            {
                0 => ClickMode.Left,
                1 => ClickMode.Right,
                _ => ClickMode.DoubleLeft,
            };

            bool randomEnabled = RandomIntervalCheckBox.IsChecked == true;
            bool useCurrentCursor = UseCurrentCursorCheckBox.IsChecked == true;
            int randomMin = intervalMs;
            int randomMax = intervalMs;

            if (!useCurrentCursor)
            {
                if (!int.TryParse(ClickXTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out x) ||
                    !int.TryParse(ClickYTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out y))
                {
                    AutoClickerStatusText.Text = "Status: vul geldige X/Y in";
                    return false;
                }
            }

            if (randomEnabled)
            {
                if (!int.TryParse(RandomMinTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out randomMin) ||
                    !int.TryParse(RandomMaxTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out randomMax) ||
                    randomMin < 1 || randomMax < randomMin)
                {
                    AutoClickerStatusText.Text = "Status: random min/max ongeldig";
                    return false;
                }
            }

            config = new AutoClickerConfig(x, y, intervalMs, clickMode, randomEnabled, randomMin, randomMax, useCurrentCursor);
            return true;
        }

        private async Task RunAutoClickerAsync(AutoClickerConfig config, CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (!config.UseCurrentCursor)
                    {
                        SetCursorPos(config.X, config.Y);
                    }

                    PerformClick(config.Mode, token);

                    int delayMs = config.UseRandomInterval
                        ? Random.Shared.Next(config.RandomMinMs, config.RandomMaxMs + 1)
                        : config.IntervalMs;

                    await Task.Delay(delayMs, token);
                }
            }
            catch (OperationCanceledException)
            {
                // Normal stop path.
            }
        }

        private void PerformClick(ClickMode mode, CancellationToken token)
        {
            switch (mode)
            {
                case ClickMode.Left:
                    mouse_event(MouseeventfLeftdown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseeventfLeftup, 0, 0, 0, UIntPtr.Zero);
                    break;
                case ClickMode.Right:
                    mouse_event(MouseeventfRightdown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseeventfRightup, 0, 0, 0, UIntPtr.Zero);
                    break;
                case ClickMode.DoubleLeft:
                    mouse_event(MouseeventfLeftdown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseeventfLeftup, 0, 0, 0, UIntPtr.Zero);
                    Task.Delay(35, token).GetAwaiter().GetResult();
                    mouse_event(MouseeventfLeftdown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseeventfLeftup, 0, 0, 0, UIntPtr.Zero);
                    break;
            }
        }

        private void SavePreset_Click(object sender, RoutedEventArgs e)
        {
            string presetName = PresetNameTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(presetName))
            {
                AutoClickerStatusText.Text = "Status: geef een preset naam op";
                return;
            }

            if (!TryReadAutoClickerConfig(out var config))
            {
                return;
            }

            var profile = CreateProfileFromUi(presetName, config);
            int existingIndex = _profiles.FindIndex(p => string.Equals(p.Name, presetName, StringComparison.OrdinalIgnoreCase));
            if (existingIndex >= 0)
            {
                _profiles[existingIndex] = profile;
            }
            else
            {
                _profiles.Add(profile);
            }

            SaveProfilesToDisk();
            RefreshPresetComboBoxItems();
            PresetComboBox.SelectedItem = presetName;
            KeyboardStatusText.Text = $"Status: preset '{presetName}' opgeslagen";
        }

        private void LoadPreset_Click(object sender, RoutedEventArgs e)
        {
            if (PresetComboBox.SelectedItem is not string selectedName)
            {
                AutoClickerStatusText.Text = "Status: kies eerst een preset";
                return;
            }

            var profile = _profiles.FirstOrDefault(p => string.Equals(p.Name, selectedName, StringComparison.OrdinalIgnoreCase));
            if (profile is null)
            {
                AutoClickerStatusText.Text = "Status: preset niet gevonden";
                return;
            }

            ApplyProfileToUi(profile);
            PresetNameTextBox.Text = profile.Name;
            KeyboardStatusText.Text = $"Status: preset '{profile.Name}' geladen";
        }

        private void DeletePreset_Click(object sender, RoutedEventArgs e)
        {
            if (PresetComboBox.SelectedItem is not string selectedName)
            {
                AutoClickerStatusText.Text = "Status: kies eerst een preset";
                return;
            }

            _profiles.RemoveAll(p => string.Equals(p.Name, selectedName, StringComparison.OrdinalIgnoreCase));
            SaveProfilesToDisk();
            RefreshPresetComboBoxItems();
            AutoClickerStatusText.Text = $"Status: preset '{selectedName}' verwijderd";
        }

        private void RefreshPresets_Click(object sender, RoutedEventArgs e)
        {
            RefreshPresets();
            AutoClickerStatusText.Text = "Status: presets ververst";
        }

        private void RefreshPresets()
        {
            _profiles = LoadProfilesFromDisk();
            RefreshPresetComboBoxItems();
        }

        private void RefreshPresetComboBoxItems()
        {
            PresetComboBox.Items.Clear();
            foreach (var profile in _profiles.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase))
            {
                PresetComboBox.Items.Add(profile.Name);
            }

            if (PresetComboBox.Items.Count > 0)
            {
                PresetComboBox.SelectedIndex = 0;
            }
        }

        private AutomationProfile CreateProfileFromUi(string presetName, AutoClickerConfig config)
        {
            return new AutomationProfile
            {
                Name = presetName,
                ClickX = config.X,
                ClickY = config.Y,
                IntervalMs = config.IntervalMs,
                ClickMode = config.Mode.ToString(),
                UseCurrentCursor = config.UseCurrentCursor,
                UseRandomInterval = config.UseRandomInterval,
                RandomMinMs = config.RandomMinMs,
                RandomMaxMs = config.RandomMaxMs,
                HotkeyKey = HotkeyComboBox.SelectedItem is Key hotkey ? hotkey.ToString() : Key.F6.ToString(),
                HotkeyCtrl = CtrlModifierCheckBox.IsChecked == true,
                HotkeyAlt = AltModifierCheckBox.IsChecked == true,
                HotkeyShift = ShiftModifierCheckBox.IsChecked == true,
                HotkeyWin = WinModifierCheckBox.IsChecked == true,
                ScriptText = ScriptTextBox.Text,
                LoopScript = LoopScriptCheckBox.IsChecked == true,
                ScriptHotkeyKey = ScriptToggleHotkeyComboBox.SelectedItem is Key scriptToggle ? scriptToggle.ToString() : Key.F7.ToString(),
                ScriptPauseHotkeyKey = ScriptPauseHotkeyComboBox.SelectedItem is Key scriptPause ? scriptPause.ToString() : Key.F8.ToString(),
                ScriptHotkeyCtrl = ScriptCtrlModifierCheckBox.IsChecked == true,
                ScriptHotkeyAlt = ScriptAltModifierCheckBox.IsChecked == true,
                ScriptHotkeyShift = ScriptShiftModifierCheckBox.IsChecked == true,
                ScriptHotkeyWin = ScriptWinModifierCheckBox.IsChecked == true,
            };
        }

        private void ApplyProfileToUi(AutomationProfile profile)
        {
            ClickXTextBox.Text = profile.ClickX.ToString(CultureInfo.InvariantCulture);
            ClickYTextBox.Text = profile.ClickY.ToString(CultureInfo.InvariantCulture);
            double seconds = Math.Max(0.001, profile.IntervalMs / 1000.0);
            SimpleIntervalSecondsTextBox.Text = seconds.ToString("0.###", CultureInfo.InvariantCulture);
            UseCurrentCursorCheckBox.IsChecked = profile.UseCurrentCursor;
            RandomIntervalCheckBox.IsChecked = profile.UseRandomInterval;
            RandomMinTextBox.Text = profile.RandomMinMs.ToString(CultureInfo.InvariantCulture);
            RandomMaxTextBox.Text = profile.RandomMaxMs.ToString(CultureInfo.InvariantCulture);

            ClickTypeComboBox.SelectedIndex = profile.ClickMode?.ToUpperInvariant() switch
            {
                "RIGHT" => 1,
                "DOUBLELEFT" => 2,
                _ => 0,
            };

            CtrlModifierCheckBox.IsChecked = profile.HotkeyCtrl;
            AltModifierCheckBox.IsChecked = profile.HotkeyAlt;
            ShiftModifierCheckBox.IsChecked = profile.HotkeyShift;
            WinModifierCheckBox.IsChecked = profile.HotkeyWin;

            if (Enum.TryParse(profile.HotkeyKey, true, out Key hotkey))
            {
                if (!HotkeyComboBox.Items.Contains(hotkey))
                {
                    HotkeyComboBox.Items.Add(hotkey);
                }

                HotkeyComboBox.SelectedItem = hotkey;
            }

            ScriptCtrlModifierCheckBox.IsChecked = profile.ScriptHotkeyCtrl;
            ScriptAltModifierCheckBox.IsChecked = profile.ScriptHotkeyAlt;
            ScriptShiftModifierCheckBox.IsChecked = profile.ScriptHotkeyShift;
            ScriptWinModifierCheckBox.IsChecked = profile.ScriptHotkeyWin;

            if (Enum.TryParse(profile.ScriptHotkeyKey, true, out Key scriptToggle))
            {
                if (!ScriptToggleHotkeyComboBox.Items.Contains(scriptToggle))
                {
                    ScriptToggleHotkeyComboBox.Items.Add(scriptToggle);
                }

                ScriptToggleHotkeyComboBox.SelectedItem = scriptToggle;
            }

            if (Enum.TryParse(profile.ScriptPauseHotkeyKey, true, out Key scriptPause))
            {
                if (!ScriptPauseHotkeyComboBox.Items.Contains(scriptPause))
                {
                    ScriptPauseHotkeyComboBox.Items.Add(scriptPause);
                }

                ScriptPauseHotkeyComboBox.SelectedItem = scriptPause;
            }

            ScriptTextBox.Text = profile.ScriptText ?? string.Empty;
            LoopScriptCheckBox.IsChecked = profile.LoopScript;
            ApplyHotkeyRegistration();
            ApplyScriptHotkeyRegistration();
        }

        private List<AutomationProfile> LoadProfilesFromDisk()
        {
            try
            {
                string path = GetProfileFilePath();
                if (!File.Exists(path))
                {
                    return new List<AutomationProfile>();
                }

                string json = File.ReadAllText(path);
                return JsonSerializer.Deserialize<List<AutomationProfile>>(json) ?? new List<AutomationProfile>();
            }
            catch
            {
                KeyboardStatusText.Text = "Status: kon presets niet laden";
                return new List<AutomationProfile>();
            }
        }

        private void SaveProfilesToDisk()
        {
            try
            {
                string path = GetProfileFilePath();
                string? folder = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                };

                string json = JsonSerializer.Serialize(_profiles, options);
                File.WriteAllText(path, json);
            }
            catch
            {
                KeyboardStatusText.Text = "Status: kon presets niet opslaan";
            }
        }

        private static string GetProfileFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "PnatClicker", "profiles.json");
        }

        private bool ApplyHotkeyRegistration()
        {
            if (_source is null)
            {
                return false;
            }

            if (HotkeyComboBox.SelectedItem is not Key selectedKey)
            {
                AutoClickerStatusText.Text = "Status: kies een key voor de keybind";
                return false;
            }

            uint modifiers = 0;
            if (CtrlModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModControl;
            }

            if (AltModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModAlt;
            }

            if (ShiftModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModShift;
            }

            if (WinModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModWin;
            }

            int virtualKey = KeyInterop.VirtualKeyFromKey(selectedKey);
            UnregisterAutoClickerHotkeyIfNeeded();

            _autoClickerHotkeyRegistered = RegisterHotKey(_source.Handle, AutoClickerHotkeyId, modifiers, (uint)virtualKey);
            if (!_autoClickerHotkeyRegistered)
            {
                AutoClickerStatusText.Text = "Status: keybind registreren mislukt";
                return false;
            }

            AutoClickerStatusText.Text = $"Status: keybind gezet ({BuildHotkeyLabel(modifiers, selectedKey)})";
            return true;
        }

        private bool ApplyScriptHotkeyRegistration()
        {
            if (_source is null)
            {
                return false;
            }

            if (ScriptToggleHotkeyComboBox.SelectedItem is not Key toggleKey ||
                ScriptPauseHotkeyComboBox.SelectedItem is not Key pauseKey)
            {
                KeyboardStatusText.Text = "Status: kies script hotkeys";
                return false;
            }

            uint modifiers = 0;
            if (ScriptCtrlModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModControl;
            }

            if (ScriptAltModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModAlt;
            }

            if (ScriptShiftModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModShift;
            }

            if (ScriptWinModifierCheckBox.IsChecked == true)
            {
                modifiers |= ModWin;
            }

            int toggleVirtualKey = KeyInterop.VirtualKeyFromKey(toggleKey);
            int pauseVirtualKey = KeyInterop.VirtualKeyFromKey(pauseKey);

            UnregisterScriptHotkeysIfNeeded();

            _scriptToggleHotkeyRegistered = RegisterHotKey(_source.Handle, ScriptToggleHotkeyId, modifiers, (uint)toggleVirtualKey);
            _scriptPauseHotkeyRegistered = RegisterHotKey(_source.Handle, ScriptPauseHotkeyId, modifiers, (uint)pauseVirtualKey);

            if (!_scriptToggleHotkeyRegistered || !_scriptPauseHotkeyRegistered)
            {
                UnregisterScriptHotkeysIfNeeded();
                KeyboardStatusText.Text = "Status: script hotkeys registreren mislukt";
                return false;
            }

            string toggleLabel = BuildHotkeyLabel(modifiers, toggleKey);
            string pauseLabel = BuildHotkeyLabel(modifiers, pauseKey);
            KeyboardStatusText.Text = $"Status: toggle={toggleLabel}, pause={pauseLabel}";
            return true;
        }

        private string BuildHotkeyLabel(uint modifiers, Key key)
        {
            var sb = new StringBuilder();

            if ((modifiers & ModControl) != 0)
            {
                sb.Append("Ctrl+");
            }

            if ((modifiers & ModAlt) != 0)
            {
                sb.Append("Alt+");
            }

            if ((modifiers & ModShift) != 0)
            {
                sb.Append("Shift+");
            }

            if ((modifiers & ModWin) != 0)
            {
                sb.Append("Win+");
            }

            sb.Append(key);
            return sb.ToString();
        }

        private void PopulateHotkeyKeys()
        {
            for (Key key = Key.F1; key <= Key.F12; key++)
            {
                HotkeyComboBox.Items.Add(key);
                ScriptToggleHotkeyComboBox.Items.Add(key);
                ScriptPauseHotkeyComboBox.Items.Add(key);
            }

            for (Key key = Key.A; key <= Key.Z; key++)
            {
                HotkeyComboBox.Items.Add(key);
                ScriptToggleHotkeyComboBox.Items.Add(key);
                ScriptPauseHotkeyComboBox.Items.Add(key);
            }
        }

        private void StopScriptPlayback()
        {
            _scriptPlaybackCts?.Cancel();
        }

        private void ToggleScriptPlaybackFromHotkey()
        {
            if (_scriptPlaybackCts is not null)
            {
                StopScriptPlayback();
                KeyboardStatusText.Text = "Status: afspelen gestopt (hotkey)";
                return;
            }

            PlayScriptButton_Click(this, new RoutedEventArgs());
        }

        private void ToggleScriptPauseFromHotkey()
        {
            if (_scriptPlaybackCts is null)
            {
                return;
            }

            _isScriptPaused = !_isScriptPaused;
            KeyboardStatusText.Text = _isScriptPaused ? "Status: script gepauzeerd" : "Status: script hervat";
        }

        private void StopRecording()
        {
            if (!_isRecording)
            {
                return;
            }

            _isRecording = false;
            _recordingStopwatch.Stop();
            RecordToggleButton.Content = "Start opname";
        }

        private bool TryParseScript(string scriptText, out List<ScriptCommand> commands, out string error)
        {
            commands = new List<ScriptCommand>();
            error = string.Empty;

            if (!scriptText.Contains("TEXT ", StringComparison.OrdinalIgnoreCase) &&
                !scriptText.Contains("WAIT ", StringComparison.OrdinalIgnoreCase) &&
                !scriptText.Contains("KEY ", StringComparison.OrdinalIgnoreCase) &&
                !scriptText.Contains("DOWN ", StringComparison.OrdinalIgnoreCase) &&
                !scriptText.Contains("UP ", StringComparison.OrdinalIgnoreCase))
            {
                string plain = scriptText.Replace("\r", string.Empty);
                if (!string.IsNullOrWhiteSpace(plain))
                {
                    commands.Add(new ScriptCommand(ScriptCommandType.Text, plain, 0, Key.None));
                }

                return true;
            }

            var lines = scriptText.Replace("\r", string.Empty).Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string raw = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                if (raw.StartsWith("TEXT ", StringComparison.OrdinalIgnoreCase))
                {
                    string text = raw[5..];
                    commands.Add(new ScriptCommand(ScriptCommandType.Text, text, 0, Key.None));
                    continue;
                }

                if (raw.StartsWith("WAIT ", StringComparison.OrdinalIgnoreCase))
                {
                    string value = raw[5..].Trim();
                    if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int waitMs) || waitMs < 0)
                    {
                        error = $"Regel {i + 1}: WAIT heeft een geldig nummer nodig.";
                        return false;
                    }

                    commands.Add(new ScriptCommand(ScriptCommandType.Wait, string.Empty, waitMs, Key.None));
                    continue;
                }

                if (raw.StartsWith("KEY ", StringComparison.OrdinalIgnoreCase) ||
                    raw.StartsWith("DOWN ", StringComparison.OrdinalIgnoreCase) ||
                    raw.StartsWith("UP ", StringComparison.OrdinalIgnoreCase))
                {
                    string[] parts = raw.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length != 2)
                    {
                        error = $"Regel {i + 1}: gebruik KEY <toets>, DOWN <toets> of UP <toets>.";
                        return false;
                    }

                    string commandName = parts[0].ToUpperInvariant();
                    if (!Enum.TryParse(parts[1], true, out Key parsedKey) || parsedKey == Key.None)
                    {
                        error = $"Regel {i + 1}: onbekende toets '{parts[1]}'.";
                        return false;
                    }

                    ScriptCommandType type = commandName switch
                    {
                        "KEY" => ScriptCommandType.Key,
                        "DOWN" => ScriptCommandType.KeyDown,
                        _ => ScriptCommandType.KeyUp,
                    };

                    commands.Add(new ScriptCommand(type, string.Empty, 0, parsedKey));
                    continue;
                }

                commands.Add(new ScriptCommand(ScriptCommandType.Text, raw, 0, Key.None));
            }

            return true;
        }

        private async Task PlayCommandsAsync(List<ScriptCommand> commands, CancellationToken token)
        {
            foreach (var command in commands)
            {
                token.ThrowIfCancellationRequested();
                await WaitIfPausedAsync(token);

                switch (command.Type)
                {
                    case ScriptCommandType.Wait:
                        await Task.Delay(command.WaitMs, token);
                        break;
                    case ScriptCommandType.Text:
                        SendUnicodeText(command.TextValue);
                        break;
                    case ScriptCommandType.Key:
                        SendKey(command.Key, true);
                        SendKey(command.Key, false);
                        break;
                    case ScriptCommandType.KeyDown:
                        SendKey(command.Key, true);
                        break;
                    case ScriptCommandType.KeyUp:
                        SendKey(command.Key, false);
                        break;
                }
            }
        }

        private async Task WaitIfPausedAsync(CancellationToken token)
        {
            while (_isScriptPaused && !token.IsCancellationRequested)
            {
                await Task.Delay(60, token);
            }
        }

        private void SendUnicodeText(string text)
        {
            foreach (char ch in text)
            {
                var inputs = new INPUT[2];

                inputs[0] = new INPUT
                {
                    type = InputKeyboard,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0,
                            wScan = ch,
                            dwFlags = KeyeventfUnicode,
                            dwExtraInfo = IntPtr.Zero,
                        },
                    },
                };

                inputs[1] = new INPUT
                {
                    type = InputKeyboard,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0,
                            wScan = ch,
                            dwFlags = KeyeventfUnicode | KeyeventfKeyup,
                            dwExtraInfo = IntPtr.Zero,
                        },
                    },
                };

                SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
            }
        }

        private void SendKey(Key key, bool isDown)
        {
            int vk = KeyInterop.VirtualKeyFromKey(key);
            if (vk <= 0)
            {
                return;
            }

            var input = new INPUT
            {
                type = InputKeyboard,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = (ushort)vk,
                        wScan = 0,
                        dwFlags = isDown ? 0 : KeyeventfKeyup,
                        dwExtraInfo = IntPtr.Zero,
                    },
                },
            };

            var buffer = new[] { input };
            SendInput((uint)buffer.Length, buffer, Marshal.SizeOf<INPUT>());
        }

        private void StartKeyboardHook()
        {
            _keyboardProc ??= KeyboardHookCallback;
            IntPtr moduleHandle = GetModuleHandle(null);
            _keyboardHookHandle = SetWindowsHookEx(WhKeyboardLl, _keyboardProc, moduleHandle, 0);
        }

        private void StopKeyboardHook()
        {
            if (_keyboardHookHandle == IntPtr.Zero)
            {
                return;
            }

            UnhookWindowsHookEx(_keyboardHookHandle);
            _keyboardHookHandle = IntPtr.Zero;
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && _isRecording)
            {
                int message = wParam.ToInt32();
                bool isDown = message == WmKeydown || message == WmSyskeydown;
                bool isUp = message == WmKeyup || message == WmSyskeyup;

                if (isDown || isUp)
                {
                    var hookData = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                    const uint llkhfInjected = 0x00000010;

                    if ((hookData.flags & llkhfInjected) == 0)
                    {
                        long now = _recordingStopwatch.ElapsedMilliseconds;
                        int delay = (int)Math.Max(0, now - _lastRecordedMs);
                        _lastRecordedMs = now;

                        Key key = KeyInterop.KeyFromVirtualKey((int)hookData.vkCode);
                        Dispatcher.Invoke(() =>
                        {
                            if (delay > 0)
                            {
                                ScriptTextBox.AppendText($"WAIT {delay}{Environment.NewLine}");
                            }

                            ScriptTextBox.AppendText($"{(isDown ? "DOWN" : "UP")} {key}{Environment.NewLine}");
                            ScriptTextBox.ScrollToEnd();
                        });
                    }
                }
            }

            return CallNextHookEx(_keyboardHookHandle, nCode, wParam, lParam);
        }

        private void UnregisterAutoClickerHotkeyIfNeeded()
        {
            if (_source is null || !_autoClickerHotkeyRegistered)
            {
                return;
            }

            UnregisterHotKey(_source.Handle, AutoClickerHotkeyId);
            _autoClickerHotkeyRegistered = false;
        }

        private void UnregisterScriptHotkeysIfNeeded()
        {
            if (_source is null)
            {
                return;
            }

            if (_scriptToggleHotkeyRegistered)
            {
                UnregisterHotKey(_source.Handle, ScriptToggleHotkeyId);
                _scriptToggleHotkeyRegistered = false;
            }

            if (_scriptPauseHotkeyRegistered)
            {
                UnregisterHotKey(_source.Handle, ScriptPauseHotkeyId);
                _scriptPauseHotkeyRegistered = false;
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WmHotkey)
            {
                int hotkeyId = wParam.ToInt32();

                if (hotkeyId == AutoClickerHotkeyId)
                {
                    ToggleAutoClicker();
                    handled = true;
                }
                else if (hotkeyId == ScriptToggleHotkeyId)
                {
                    ToggleScriptPlaybackFromHotkey();
                    handled = true;
                }
                else if (hotkeyId == ScriptPauseHotkeyId)
                {
                    ToggleScriptPauseFromHotkey();
                    handled = true;
                }
            }

            return IntPtr.Zero;
        }

        private readonly record struct AutoClickerConfig(
            int X,
            int Y,
            int IntervalMs,
            ClickMode Mode,
            bool UseRandomInterval,
            int RandomMinMs,
            int RandomMaxMs,
            bool UseCurrentCursor);

        private enum ScriptCommandType
        {
            Wait,
            Text,
            Key,
            KeyDown,
            KeyUp,
        }

        private readonly record struct ScriptCommand(ScriptCommandType Type, string TextValue, int WaitMs, Key Key);

        private sealed class AutomationProfile
        {
            public string Name { get; set; } = "Default";
            public int ClickX { get; set; } = 500;
            public int ClickY { get; set; } = 500;
            public int IntervalMs { get; set; } = 100;
            public string ClickMode { get; set; } = "Left";
            public bool UseCurrentCursor { get; set; } = true;
            public bool UseRandomInterval { get; set; }
            public int RandomMinMs { get; set; } = 80;
            public int RandomMaxMs { get; set; } = 130;
            public string HotkeyKey { get; set; } = "F6";
            public bool HotkeyCtrl { get; set; } = true;
            public bool HotkeyAlt { get; set; }
            public bool HotkeyShift { get; set; }
            public bool HotkeyWin { get; set; }
            public string ScriptText { get; set; } = "";
            public bool LoopScript { get; set; }
            public string ScriptHotkeyKey { get; set; } = "F7";
            public string ScriptPauseHotkeyKey { get; set; } = "F8";
            public bool ScriptHotkeyCtrl { get; set; } = true;
            public bool ScriptHotkeyAlt { get; set; }
            public bool ScriptHotkeyShift { get; set; }
            public bool ScriptHotkeyWin { get; set; }
        }

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;

            [FieldOffset(0)]
            public KEYBDINPUT ki;
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
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string? lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
    }
}