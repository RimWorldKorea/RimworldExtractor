namespace RimworldExtractorInternal.Compats
{
    public class CompatPriorityAttribute : Attribute
    {
        public int Priority { get; private set; }

        public CompatPriorityAttribute(int priority)
        {
            Priority = priority;
        }

        public CompatPriorityAttribute() : this(100) {}
    }
}