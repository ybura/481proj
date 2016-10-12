using System.Speech.Synthesis;
using KeePass.Plugins;
using KeePass.Forms;
using KeePassLib;

namespace keepassPlugin
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
            PwEntry current_entry;
            synthesizer.Speak("Foo");

            return true;
        }
    }
}
