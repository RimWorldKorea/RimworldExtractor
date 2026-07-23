using System.Xml.Linq;
using RimworldExtractorInternal.DataTypes;

namespace RimworldExtractorInternal.Compats
{
    public abstract class BaseCompat
    {
        public virtual IEnumerable<TranslationEntry> DoPostProcessing(IEnumerable<TranslationEntry> entries)
        {
            foreach (var entry in entries)
            {
                yield return entry;
            }
        }

        public virtual void DoPreProcessing(XDocument doc)
        {
            return;
        }
    }
}