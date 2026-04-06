using System.Diagnostics;
using System.Media;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace win9xplorer
{
    internal sealed class VolumePopupForm : Form
    {
        [DllImport("winmm.dll")]
        private static extern int waveOutGetVolume(IntPtr hwo, out uint pdwVolume);

        [DllImport("winmm.dll")]
        private static extern int waveOutSetVolume(IntPtr hwo, uint dwVolume);

        [DllImport("winmm.dll")]
        private static extern int waveOutGetNumDevs();

        [DllImport("winmm.dll", CharSet = CharSet.Unicode)]
        private static extern int waveOutGetDevCaps(IntPtr uDeviceID, out WAVEOUTCAPS pwoc, uint cbwoc);

        [DllImport("winmm.dll")]
        private static extern int waveInGetNumDevs();

        [DllImport("winmm.dll", CharSet = CharSet.Unicode)]
        private static extern int waveInGetDevCaps(IntPtr uDeviceID, out WAVEINCAPS pwic, uint cbwic);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WAVEOUTCAPS
        {
            public ushort wMid;
            public ushort wPid;
            public uint vDriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szPname;
            public uint dwFormats;
            public ushort wChannels;
            public ushort wReserved1;
            public uint dwSupport;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WAVEINCAPS
        {
            public ushort wMid;
            public ushort wPid;
            public uint vDriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szPname;
            public uint dwFormats;
            public ushort wChannels;
            public ushort wReserved1;
        }

        private sealed record AudioDeviceItem(int DeviceId, string Name)
        {
            public override string ToString() => Name;
        }

        private readonly TrackBar volumeTrack;
        private readonly CheckBox muteBox;
        private readonly ComboBox outputDeviceCombo;
        private readonly ComboBox inputDeviceCombo;
        private bool suppressEvents;
        private int lastNonZeroVolume = 70;
        private int selectedOutputDeviceId = 0;
        private bool isDraggingVolumeTrack;
        private bool volumeChangedWhileDragging;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public bool PlayFeedbackSoundOnMouseUp { get; set; } = true;

        public VolumePopupForm()
        {
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Text = "Volume Control";
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point);
            ClientSize = new Size(210, 320);
            BackColor = SystemColors.Control;

            var outputLabel = new Label
            {
                Text = "Output:",
                Left = 10,
                Top = 10,
                Width = 50
            };

            outputDeviceCombo = new ComboBox
            {
                Left = 62,
                Top = 6,
                Width = 136,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            outputDeviceCombo.SelectedIndexChanged += OutputDeviceCombo_SelectedIndexChanged;

            var inputLabel = new Label
            {
                Text = "Input:",
                Left = 10,
                Top = 34,
                Width = 50
            };

            inputDeviceCombo = new ComboBox
            {
                Left = 62,
                Top = 30,
                Width = 136,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            var channelPanel = new Panel
            {
                Left = 12,
                Top = 58,
                Width = 92,
                Height = 248,
                BorderStyle = BorderStyle.Fixed3D,
                BackColor = SystemColors.Control
            };

            var channelLabel = new Label
            {
                Text = "Volume Control",
                Width = 84,
                Height = 28,
                Left = 2,
                Top = 4,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = Font
            };

            volumeTrack = new TrackBar
            {
                Orientation = Orientation.Vertical,
                Minimum = 0,
                Maximum = 100,
                TickStyle = TickStyle.None,
                SmallChange = 2,
                LargeChange = 10,
                Height = 162,
                Width = 40,
                Left = 24,
                Top = 34
            };
            volumeTrack.ValueChanged += VolumeTrack_ValueChanged;
            volumeTrack.MouseDown += VolumeTrack_MouseDown;
            volumeTrack.MouseUp += VolumeTrack_MouseUp;

            muteBox = new CheckBox
            {
                Text = "Mute",
                AutoSize = true,
                Left = 20,
                Top = 206,
                FlatStyle = FlatStyle.Standard
            };
            muteBox.CheckedChanged += MuteBox_CheckedChanged;

            var mixerButton = new Button
            {
                Text = "Mixer...",
                Width = 86,
                Height = 22,
                Left = 114,
                Top = 58,
                FlatStyle = FlatStyle.Standard
            };
            mixerButton.Click += (_, _) => OpenMixer();

            var closeButton = new Button
            {
                Text = "Close",
                Width = 86,
                Height = 22,
                Left = 114,
                Top = 86,
                FlatStyle = FlatStyle.Standard
            };
            closeButton.Click += (_, _) => Hide();

            Controls.Add(outputLabel);
            Controls.Add(outputDeviceCombo);
            Controls.Add(inputLabel);
            Controls.Add(inputDeviceCombo);
            channelPanel.Controls.Add(channelLabel);
            channelPanel.Controls.Add(volumeTrack);
            channelPanel.Controls.Add(muteBox);
            Controls.Add(channelPanel);
            Controls.Add(mixerButton);
            Controls.Add(closeButton);

            Deactivate += (_, _) => Hide();
            Shown += (_, _) => RefreshFromSystemVolume();
            FormClosing += VolumePopupForm_FormClosing;
        }

        private void VolumePopupForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }

        public void ShowAt(Point location)
        {
            Location = location;
            RefreshAudioDevices();
            RefreshFromSystemVolume();
            Show();
            Activate();
        }

        private void RefreshAudioDevices()
        {
            suppressEvents = true;

            var outputDevices = GetOutputDevices();
            outputDeviceCombo.Items.Clear();
            foreach (var output in outputDevices)
            {
                outputDeviceCombo.Items.Add(output);
            }

            var outputSelection = outputDevices.FindIndex(x => x.DeviceId == selectedOutputDeviceId);
            outputDeviceCombo.SelectedIndex = outputSelection >= 0 ? outputSelection : (outputDevices.Count > 0 ? 0 : -1);

            var inputDevices = GetInputDevices();
            inputDeviceCombo.Items.Clear();
            foreach (var input in inputDevices)
            {
                inputDeviceCombo.Items.Add(input);
            }

            inputDeviceCombo.SelectedIndex = inputDevices.Count > 0 ? 0 : -1;
            suppressEvents = false;
        }

        private static List<AudioDeviceItem> GetOutputDevices()
        {
            var devices = new List<AudioDeviceItem>();
            var count = waveOutGetNumDevs();

            for (var i = 0; i < count; i++)
            {
                if (waveOutGetDevCaps(new IntPtr(i), out var caps, (uint)Marshal.SizeOf<WAVEOUTCAPS>()) == 0)
                {
                    devices.Add(new AudioDeviceItem(i, caps.szPname));
                }
            }

            return devices;
        }

        private static List<AudioDeviceItem> GetInputDevices()
        {
            var devices = new List<AudioDeviceItem>();
            var count = waveInGetNumDevs();

            for (var i = 0; i < count; i++)
            {
                if (waveInGetDevCaps(new IntPtr(i), out var caps, (uint)Marshal.SizeOf<WAVEINCAPS>()) == 0)
                {
                    devices.Add(new AudioDeviceItem(i, caps.szPname));
                }
            }

            return devices;
        }

        private void OutputDeviceCombo_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (suppressEvents)
            {
                return;
            }

            if (outputDeviceCombo.SelectedItem is AudioDeviceItem device)
            {
                selectedOutputDeviceId = device.DeviceId;
                RefreshFromSystemVolume();
            }
        }

        private void RefreshFromSystemVolume()
        {
            suppressEvents = true;
            var current = GetSystemVolume(selectedOutputDeviceId);
            volumeTrack.Value = current;
            muteBox.Checked = current == 0;
            if (current > 0)
            {
                lastNonZeroVolume = current;
            }

            suppressEvents = false;
        }

        private void VolumeTrack_ValueChanged(object? sender, EventArgs e)
        {
            if (suppressEvents)
            {
                return;
            }

            var value = volumeTrack.Value;
            SetSystemVolume(selectedOutputDeviceId, value);

            suppressEvents = true;
            muteBox.Checked = value == 0;
            suppressEvents = false;

            if (value > 0)
            {
                lastNonZeroVolume = value;
            }

            if (isDraggingVolumeTrack)
            {
                volumeChangedWhileDragging = true;
            }
        }

        private void VolumeTrack_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            isDraggingVolumeTrack = true;
            volumeChangedWhileDragging = false;
        }

        private void VolumeTrack_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            if (volumeChangedWhileDragging && PlayFeedbackSoundOnMouseUp)
            {
                PlayVolumeFeedbackSound();
            }

            isDraggingVolumeTrack = false;
            volumeChangedWhileDragging = false;
        }

        private void MuteBox_CheckedChanged(object? sender, EventArgs e)
        {
            if (suppressEvents)
            {
                return;
            }

            if (muteBox.Checked)
            {
                if (volumeTrack.Value > 0)
                {
                    lastNonZeroVolume = volumeTrack.Value;
                }

                volumeTrack.Value = 0;
                SetSystemVolume(selectedOutputDeviceId, 0);
            }
            else
            {
                var restored = Math.Max(5, lastNonZeroVolume);
                volumeTrack.Value = restored;
                SetSystemVolume(selectedOutputDeviceId, restored);
            }
        }

        private static int GetSystemVolume(int deviceId)
        {
            if (waveOutGetVolume(new IntPtr(deviceId), out var volume) != 0)
            {
                return 50;
            }

            var left = (int)(volume & 0xFFFF);
            return Math.Clamp((left * 100) / 0xFFFF, 0, 100);
        }

        private static void SetSystemVolume(int deviceId, int value)
        {
            value = Math.Clamp(value, 0, 100);
            var scaled = (uint)((value * 0xFFFF) / 100);
            var bothChannels = (scaled & 0xFFFF) | (scaled << 16);
            waveOutSetVolume(new IntPtr(deviceId), bothChannels);
        }

        private static void OpenMixer()
        {
            var candidates = new[]
            {
                "sndvol32.exe",
                "sndvol.exe",
                "control.exe"
            };

            foreach (var candidate in candidates)
            {
                try
                {
                    var info = new ProcessStartInfo
                    {
                        FileName = candidate,
                        Arguments = candidate == "control.exe" ? "mmsys.cpl" : string.Empty,
                        UseShellExecute = true
                    };

                    Process.Start(info);
                    return;
                }
                catch
                {
                }
            }
        }

        private static void PlayVolumeFeedbackSound()
        {
            try
            {
                SystemSounds.Asterisk.Play();
            }
            catch
            {
            }
        }
    }
}
