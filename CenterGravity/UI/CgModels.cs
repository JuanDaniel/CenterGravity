namespace BBI.JD.UI
{
    /// <summary>Where the reported X/Y/Z are measured from.</summary>
    public enum ReferenceMode
    {
        InternalOrigin = 0,
        ProjectBasePoint = 1,
        SurveyPoint = 2
    }

    /// <summary>How the combined centre of gravity is weighted.</summary>
    public enum WeightMode
    {
        Volume = 0,
        Mass = 1
    }

    /// <summary>One selected element as shown in the pane grid (display strings only).</summary>
    public class CgRow
    {
        public string Id { get; set; }
        public string Category { get; set; }
        public string Name { get; set; }

        /// <summary>Volume ready for display (unit symbol included), or "-" when skipped.</summary>
        public string Volume { get; set; }

        /// <summary>Volume in Revit internal units (cubic feet); 0 when skipped.</summary>
        public double VolumeInternal { get; set; }

        /// <summary>Mass ready for display, or "-" when unknown / skipped.</summary>
        public string Mass { get; set; }

        /// <summary>Mass in Revit internal units; 0 when unknown.</summary>
        public double MassInternal { get; set; }

        /// <summary>True when a real mass (from material density) was found for this element.</summary>
        public bool MassKnown { get; set; }

        /// <summary>True when the element carried no usable solid and was left out of the calculation.</summary>
        public bool Skipped { get; set; }
    }
}
