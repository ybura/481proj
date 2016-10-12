using System;
using System.Collections.Generic;
using System.Speech.Synthesis;

using KeePass.Plugins;
using KeePass.Forms.MainForm;

namespace OurPlugin
{
	public sealed class SamplePluginExt : Plugin
	{
		private IPluginHost m_host = null;

		public override bool Initialize(IPluginHost host)
		{
			m_host = host;
			SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = -2;     // -10...10

            // Synchronous
            PwGroup currentGroup = GetSelectedGroup()
            synthesizer.Speak(currentGroup);

			return true;
		}
	}
}
