namespace Toppling
{  
    public class FormParameters
    {
        public MainFormParameters? MainFormParameters { get; set; }
        public AddSupportParameters? AddSupportParameters { get; set; }
    }
    public class MainFormParameters
    {
        public double SlopeHeight { get; set; }
        public double SlopeAngle { get; set; }
        public double TopAngle { get; set; }
        public double SlopeDipDir { get; set; }
        public double MeanDipA { get; set; }
        public double MeanDipB { get; set; }
        public double MeanDipDirA { get; set; }
        public double MeanDipDirB { get; set; }
        public double MeanSpaceA { get; set; }
        public double MeanSpaceB { get; set; }
        public double MeanFricA { get; set; }
        public double MeanFricB { get; set; }
        public double MeanSeis { get; set; }
        public double PorePress { get; set; }
        public double UnitWeight { get; set; }
        public double UnitWeightH2O { get; set; }
        public string? DistSpaceA { get; set; }
        public string? DistSpaceB { get; set; }
        public string? DistFricA { get; set; }
        public string? DistFricB { get; set; }
        public string? DistSeis { get; set; }
        public string? DistPorePress { get; set; }
        public string? DistUnitWeight { get; set; }
        public string? DistUnitWeightH2O { get; set; }
        public double FisherKA { get; set; }
        public double FisherKB { get; set; }
        public double StDevSpaceA { get; set; }
        public double StDevSpaceB { get; set; }
        public double StDevFricA { get; set; }
        public double StDevFricB { get; set; }
        public double StDevUnitWt { get; set; }
        public double StDevUnitWtH20 { get; set; }
        public double StDevSeis { get; set; }
        public string? StDevPorePress { get; set; }
        public bool BoltBlocksTogether { get; set; }
        public bool AddToeSupport { get; set; }
    }

    public class AddSupportParameters
    {
        public string? NatureOfForceApplication { get; set; }
        public double Magnitude { get; set; }
        public double Orientation { get; set; }
        public double OptimumOrientationAgainstSliding { get; set; }
        public double OptimumOrientationAgainstToppling { get; set; }
        public double EffectiveWidth { get; set; }
    }
}
