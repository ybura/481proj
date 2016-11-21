using System.Speech.Synthesis;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using KeePass.Plugins;
using KeePassLib;
using KeePass.UI;
using KeePass;

namespace IndiaPlugin
{
    public sealed class IndiaPluginExt : Plugin
    {
        private IPluginHost m_host = null;

        private ToolStripSeparator m_tsSeparator = null;
        private ToolStripMenuItem m_tsmiPopup = null;        

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // On creation of plugin, set-up resources.
        public override bool Initialize(IPluginHost host)
        {
            Terminate();
            if (host == null) return false;
            m_host = host;

            m_host.MainWindow.UIStateUpdated += this.OnUIStateUpdated;
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            synthesizer.Speak("Hi India.");

            ToolStripItemCollection tsMenu = m_host.MainWindow.ToolsMenu.DropDownItems;

            m_host.MainWindow.KeyPreview = true;
            m_host.MainWindow.KeyDown += SpeakSelectedEntry;
            m_host.MainWindow.KeyDown += SpeakAllEntries;
            m_host.MainWindow.KeyDown += HelpMenu;
            // Add a separator at the bottom
            m_tsSeparator = new ToolStripSeparator();
            tsMenu.Add(m_tsSeparator);

            // Add the popup menu item
            m_tsmiPopup = new ToolStripMenuItem();
            m_tsmiPopup.Text = "Plugin for India";
            tsMenu.Add(m_tsmiPopup);
            
            // TODO: Set-up hotkey options for plugin.
            //m_hkGlobalAutoType.HotKey = (kAT & Keys.KeyCode);
            //m_hkGlobalAutoType.HotKeyModifiers = (kAT & Keys.Modifiers);
            //m_hkGlobalAutoType.RenderHotKey();
           
            return true;
       } 

        // Destruction of plugin, clean-up resources.
        public override void Terminate()
        {
            if (m_host == null) return;

            m_host.MainWindow.UIStateUpdated -= this.OnUIStateUpdated;

            m_host = null;
        }
        
        private void HelpMenu(object sender, KeyEventArgs e)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10

            if (e.KeyCode == Keys.H && e.Control)
            {
                synthesizer.Speak("These are the keyboard options: press up or down to highlight an entry and have it spoken out, press control E to hear all the entries in the database, press control U to open the URL of the selected entry, press control H to hear this help menu again.");
            }
        }

        private void SpeakSelectedEntry(object sender, KeyEventArgs e)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            if (!m_host.Database.IsOpen)
            {
                synthesizer.Speak("You first need to open a database!");
                return;
            }

            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);
            System.Windows.Forms.ListView.SelectedIndexCollection selected_index = lv.SelectedIndices;

            if (e.KeyCode == Keys.Up)
            {                
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    if (selected_index.Contains(i))
                    {
                        if (i - 1 < 0)
                        {
                            synthesizer.Speak(lv.SelectedItems[0].Text);
                        }
                        else
                        {
                            synthesizer.Speak(lv.Items[i - 1].Text);
                        }
                    }
                }
            }
            if (e.KeyCode == Keys.Down)
            {
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    if (selected_index.Contains(i))
                    {
                        if (i + 1 >= lv.Items.Count)
                        {
                            synthesizer.Speak(lv.SelectedItems[0].Text);
                        }
                        else
                        {
                            synthesizer.Speak(lv.Items[i + 1].Text);
                        }
                    }
                }
            }
        }
        
        // Dictates all the items in the password listview.
        private void SpeakAllEntries(object sender, KeyEventArgs e)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            
            if (!m_host.Database.IsOpen)
            {
                synthesizer.Speak("You first need to open a database!");
                return;
            }

            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            if (e.KeyCode == Keys.E && e.Control)
            {
                synthesizer.Speak("I am about to tell you the current entries.");
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    synthesizer.Speak(lv.Items[i].Text);
                }
            }           

        }
        
        private void OpenURL(object sender, KeyEventArgs e)
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            if (e.KeyCode == Keys.U && e.Control)
            {
                //Open url
                System.Diagnostics.Process.Start(lv.SelectedItems[3].Text);
            }
        }
        private void SetSelectedEntryType(object sender, EventArgs e)
        // Dictates the tag for the currently selected entry in the listview.
        // TODO: Does not work as expected.
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            int index = lv.FocusedItem.Index;
        }

        // UI Refresh callback function.
        private void OnUIStateUpdated(object sender, EventArgs e)
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            lv.BeginUpdate();
            /*
            for (int i = 0; i < lv.Items.Count; i++)
            {
                lv.Items[i].Font = new Font("Arial", 36);
            }*/
            for (int i = 0; i < lv.Items.Count; i++)
            {
                if (i % 2 == 0)
                {
                    lv.Items[i].BackColor = System.Drawing.Color.DarkSalmon;
                }
                else if (i % 2 == 1)
                {
                    lv.Items[i].BackColor = System.Drawing.Color.Blue;
                }
            }

            lv.EndUpdate();
        }
        
    }
}
