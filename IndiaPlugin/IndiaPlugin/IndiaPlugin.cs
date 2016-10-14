using System.Speech.Synthesis;
using KeePass.Plugins;
using KeePassLib;

namespace IndiaPlugin
{
    public sealed class IndiaPluginExt : Plugin
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
            synthesizer.Speak("Hello");

            return true;
        }
        public override void Terminate()
        {
        }
    }
}
