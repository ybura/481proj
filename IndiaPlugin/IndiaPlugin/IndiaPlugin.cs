using System.Speech.Synthesis;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using KeePass.Plugins;
using KeePassLib;
using KeePass.UI;

namespace IndiaPlugin
{
    public sealed class IndiaPluginExt : Plugin
    {
        private IPluginHost m_host = null;

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        public override bool Initialize(IPluginHost host)
        {
            Terminate();

            if (host == null) return false;
            m_host = host;

            m_host.MainWindow.UIStateUpdated += this.OnUIStateUpdated;

            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = -2;     // -10...10
            synthesizer.Speak("Hi India.");

            return true;
        }
        public override void Terminate()
        {
            if (m_host == null) return;

            m_host.MainWindow.UIStateUpdated -= this.OnUIStateUpdated;

            m_host = null;
        }

        private void OnUIStateUpdated(object sender, EventArgs e)
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            lv.BeginUpdate();
            for (int i = 0; i < lv.Items.Count; i++)
            {
                lv.Items[i].Font = new Font("Arial", 36);
            }
            lv.EndUpdate();
        }
        
    }
}
