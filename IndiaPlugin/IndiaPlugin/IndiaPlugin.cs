﻿using System.Speech.Synthesis;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using KeePass.UI;
using KeePass;


using KeePass.Plugins;
using KeePass.Forms;

namespace IndiaPlugin
{
    public sealed class IndiaPluginExt : Plugin
    {
        private IPluginHost m_host = null;
        private EventHandler<GwmWindowEventArgs> global_window_manager_window_added_handler;
        private int tabCount = 0;
        private PwEntryForm form;
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();
        VoiceAge speechAge = VoiceAge.Adult;
        VoiceGender speechGender = VoiceGender.Female;

        bool selected = false;
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        
        // On creation of plugin, set-up resources.
        public override bool Initialize(IPluginHost host)
        {
            Terminate();
            if (host == null) return false;
            m_host = host;

            global_window_manager_window_added_handler = new EventHandler<GwmWindowEventArgs>(GlobalWindowManager_WindowAdded);
            GlobalWindowManager.WindowAdded += global_window_manager_window_added_handler;
            m_host.MainWindow.UIStateUpdated += this.OnUIStateUpdated;

            synthesizer.SelectVoiceByHints(speechGender, speechAge);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            synthesizer.SpeakAsync("Hi India. Type your password and press enter.");
            

            m_host.MainWindow.KeyPreview = true;
            m_host.MainWindow.KeyDown += SpeakSelectedEntry;
            m_host.MainWindow.KeyDown += SpeakAllEntries;
            m_host.MainWindow.KeyDown += HelpMenu;
            m_host.MainWindow.KeyDown += KillSpeech;
            m_host.MainWindow.KeyDown += CustomizeVoice;
            m_host.MainWindow.KeyDown += SelectFemaleVoice;
            m_host.MainWindow.KeyDown += SelectMaleVoice;

            return true;
       } 

        // Destruction of plugin, clean-up resources.
        public override void Terminate()
        {
            if (m_host == null) return;

            m_host.MainWindow.UIStateUpdated -= this.OnUIStateUpdated;
            m_host.MainWindow.KeyDown -= SpeakSelectedEntry;
            m_host.MainWindow.KeyDown -= SpeakAllEntries;
            m_host.MainWindow.KeyDown -= HelpMenu;
            m_host.MainWindow.KeyDown -= KillSpeech;
            m_host.MainWindow.KeyDown -= CustomizeVoice;
            m_host.MainWindow.KeyDown -= SelectFemaleVoice;
            m_host.MainWindow.KeyDown -= SelectMaleVoice;
            GlobalWindowManager.WindowAdded -= global_window_manager_window_added_handler;
            m_host = null;
        }

        void GlobalWindowManager_WindowAdded(object sender, GwmWindowEventArgs e)
        {
            form = e.Form as PwEntryForm;

            synthesizer.SelectVoiceByHints(speechGender, speechAge);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            
            if (form== null)
                return;

            form.KeyPreview = true;
            form.KeyDown += KillSpeech;
            form.KeyDown += SpeakAddEntryMenu;
            form.KeyDown += AddHelpMenu;
        }
        
        private void SpeakAddEntryMenu(object sender, KeyEventArgs e)
        {
            synthesizer.SelectVoiceByHints(speechGender, speechAge);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10

            TextBox m_tbUserName = (form.Controls.Find("m_tbUserName", true)[0] as TextBox);
            TextBox m_tbUrl= (form.Controls.Find("m_tbUrl", true)[0] as TextBox);
            TextBox m_tbTitle = (form.Controls.Find("m_tbTitle", true)[0] as TextBox);
            RichTextBox m_rtNotes = (form.Controls.Find("m_rtNotes", true)[0] as RichTextBox);
            Button m_btnOk = (form.Controls.Find("m_btnOk", true)[0] as Button);

            if (e.KeyCode == Keys.N && e.Control)
            {
                if (string.IsNullOrWhiteSpace(m_tbTitle.Text))
                {
                    m_tbTitle.Focus();
                    synthesizer.SpeakAsync("Type in the title of the entry then press control N ");
                    return;
                }

                if (string.IsNullOrWhiteSpace(m_tbUserName.Text))
                {
                    m_tbUserName.Focus();
                    synthesizer.SpeakAsync("Type in your username then press control N");
                    return;
                }

                if (string.IsNullOrWhiteSpace(m_tbUrl.Text))
                {
                    m_tbUrl.Focus();
                    synthesizer.SpeakAsync("Type in the url of the website then press control N. Make sure the url is in the form of www.URL.com");
                    return;
                }

                if (string.IsNullOrWhiteSpace(m_rtNotes.Text) && tabCount < 1)
                {
                    tabCount += 1;
                    m_rtNotes.Focus();
                    synthesizer.SpeakAsync("Optional: type in some notes and press enter.");
                    return;
                }

                else
                {
                    m_btnOk.Focus();
                    synthesizer.SpeakAsync("Press enter to create your entry");
                    return;
                }
            }
        }

        private void AddHelpMenu(object sender, KeyEventArgs e)
        {
            synthesizer.SelectVoiceByHints(speechGender, speechAge);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10

            if (e.KeyCode == Keys.H && e.Control)
            {
                synthesizer.SpeakAsync("This is the field to add a new account entry. You need to have the title, username, and url form fields filled in before you can finish adding the entry. Press control N to hear which entry you need to add.");
            }
        }

        private void HelpMenu(object sender, KeyEventArgs e)
        {
            synthesizer.SelectVoiceByHints(speechGender, speechAge);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10

            if (e.KeyCode == Keys.M && e.Control)
            {
                synthesizer.SpeakAsync("These are the keyboard options: press up or down to highlight an entry and have it spoken out, press control E to hear all the entries in the database, press control I to add a new entry, press control U to open the URL of the selected entry,press control H to hear this help menu again.");
            }
        }

        private void KillSpeech(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.K && e.Control) || e.KeyCode == Keys.Enter)
            {
                synthesizer.SpeakAsyncCancelAll();
            }
        }

        private void SelectFemaleVoice(object sender, KeyEventArgs e)
        {
            if((e.KeyCode == Keys.D1 && e.Control))
            {
                speechGender = VoiceGender.Female;
                speechAge = VoiceAge.Adult;
                synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
                synthesizer.Speak("You have selected this voice");
            }
        }

        private void SelectMaleVoice(object sender, KeyEventArgs e)
        {
            if((e.KeyCode == Keys.D2 && e.Control))
            {
                speechGender = VoiceGender.Male;
                speechAge = VoiceAge.Adult;
                synthesizer.SelectVoiceByHints(VoiceGender.Male, VoiceAge.Adult);
                synthesizer.Speak("You have selected this voice");
            }
        }

        private void CustomizeVoice(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Z && e.Control))
            {
                synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
                synthesizer.Volume = 100;  // 0...100
                synthesizer.Rate = 0;     // -10...10
                synthesizer.Speak("Press control 1 to select this voice.");
                
                synthesizer.SelectVoiceByHints(VoiceGender.Male, VoiceAge.Adult);
                synthesizer.Volume = 100;  // 0...100
                synthesizer.Rate = 0;     // -10...10
                synthesizer.Speak("Press control 2 to select this voice.");
            }
        }
        
        private void SpeakSelectedEntry(object sender, KeyEventArgs e)
        {
            synthesizer.SelectVoiceByHints(speechGender, speechAge);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            if (!m_host.Database.IsOpen)
            {
                synthesizer.SpeakAsync("You first need to open a database!");
                return;
            }

            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);
            System.Windows.Forms.ListView.SelectedIndexCollection selected_index = lv.SelectedIndices;

            if (e.KeyCode == Keys.Up)
            {
                synthesizer.SpeakAsyncCancelAll();
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    if (selected_index.Contains(i))
                    {
                        if (i - 1 < 0)
                        {
                            synthesizer.SpeakAsync(lv.SelectedItems[0].Text);
                        }
                        else
                        {
                            synthesizer.SpeakAsync(lv.Items[i - 1].Text);
                        }
                    }
                }
            }
            if (e.KeyCode == Keys.Down)
            {
                synthesizer.SpeakAsyncCancelAll();
                if (lv.Items.Count == 0)
                {
                    lv.Focus();
                    lv.Items[0].Selected = true;
                }
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    if (selected_index.Contains(i))
                    {
                        if (i + 1 >= lv.Items.Count)
                        {
                            synthesizer.SpeakAsync(lv.SelectedItems[0].Text);
                        }
                        else
                        {
                            synthesizer.SpeakAsync(lv.Items[i + 1].Text);
                        }
                    }
                }
            }
        }
        
        // Dictates all the items in the password listview.
        private void SpeakAllEntries(object sender, KeyEventArgs e)
        {
             
            synthesizer.SelectVoiceByHints(speechGender, speechAge);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            
            if (!m_host.Database.IsOpen)
            {
                synthesizer.SpeakAsync("You first need to open a database!");
                return;
            }

            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            if (e.KeyCode == Keys.E && e.Control)
            {
                synthesizer.SpeakAsync("I am about to tell you the current entries.");
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    synthesizer.SpeakAsync(lv.Items[i].Text);
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

        // UI Refresh callback function.
        private void OnUIStateUpdated(object sender, EventArgs e)
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            lv.BeginUpdate();

            for (int i = 0; i < lv.Items.Count; i++)
            {
                if (i % 2 == 0)
                {
                    lv.Items[i].BackColor = System.Drawing.Color.DarkSalmon;
                }
                else if (i % 2 == 1)
                {
                    lv.Items[i].BackColor = System.Drawing.Color.Blue;
                    lv.Items[i].ForeColor = System.Drawing.Color.WhiteSmoke;
                }
            }

            lv.EndUpdate();
        }
        
    }
}
