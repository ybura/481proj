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

        private ToolStripSeparator m_tsSeparator = null;
        private ToolStripMenuItem m_tsmiPopup = null;
        private ToolStripMenuItem m_tsmiAddGroups = null;
        private ToolStripMenuItem m_tsmiAddEntries = null;

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

            ToolStripItemCollection tsMenu = m_host.MainWindow.ToolsMenu.DropDownItems;

            // Add a separator at the bottom
            m_tsSeparator = new ToolStripSeparator();
            tsMenu.Add(m_tsSeparator);

            // Add the popup menu item
            m_tsmiPopup = new ToolStripMenuItem();
            m_tsmiPopup.Text = "Plugin for India";
            tsMenu.Add(m_tsmiPopup);

            // Add menu item 'Add Some Groups'
            m_tsmiAddGroups = new ToolStripMenuItem();
            m_tsmiAddGroups.Text = "Listen to Current Entries";
            m_tsmiAddGroups.Click += SpeakEntriesMenuItem;
            m_tsmiPopup.DropDownItems.Add(m_tsmiAddGroups);
            m_tsmiPopup.ShortcutKeys = KeysControl | Keys.P;
            
            return true;
        }
        public override void Terminate()
        {
            if (m_host == null) return;

            m_host.MainWindow.UIStateUpdated -= this.OnUIStateUpdated;

            m_host = null;
        }

        private void SpeakEntriesMenuItem(object sender, EventArgs e)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = -2;     // -10...10
            if (!m_host.Database.IsOpen)
            {
                synthesizer.Speak("You first need to open a database!");
                return;
            }

            synthesizer.Speak("I am about to tell you the current entries.");
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            for (int i = 0; i < lv.Items.Count; i++)
            {
                synthesizer.Speak(lv.Items[i].Text);
            }

        }
        private void SpeakSelectedEntry(object sender, EventArgs e)
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            int index = lv.FocusedItem.Index;
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
