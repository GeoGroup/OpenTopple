namespace Toppling
{
    public class GlobalParameter
    {
        public double SlopeHeight { get; set; }     //H
        public double SlopeAngle { get; set; }      //ψs
        public double TopAngle { get; set; }        //ψts
        public double SlopeDipDir { get; set; }     //αs
        public double MeanDipA { get; set; }        //ψa
        public double MeanDipB { get; set; }        //ψb
        public double MeanDipDirA { get; set; }     //αa
        public double MeanDipDirB { get; set; }     //αb
        public double MeanSpaceA { get; set; }      //Sa
        public double MeanSpaceB { get; set; }      //Sb
        public double MeanFricA { get; set; }       //φa
        public double MeanFricB { get; set; }       //φa   
        public double UnitWeight { get; set; }  //γrock
        public double UnitWeightH2O { get; set; }   //γwater
        public double MeanSeis { get; set; }        //0.25g
        public double PorePress { get; set; }       //0.25
        public int NumTrials { get; set; }           //1000

        public string DistSpaceA { get; set; }
        public string DistSpaceB { get; set; }
        public string DistFricA { get; set; }
        public string DistFricB { get; set; }
        public string DistSeis { get; set; }
        public string DistPorePress { get; set; }
        public string DistUnitWeight { get; set; }
        public string DistUnitWeightH2O { get; set; }
        public double FisherKA { get; set; }
        public double FisherKB { get; set; }
        public double StDevSpaceA { get; set; }
        public double StDevSpaceB { get; set; }
        public double StDevFricA { get; set; }
        public double StDevFricB { get; set; }
        public double StDevSeis { get; set; }
        public double StDevPorePress { get; set; }
        public double StDevUnitWt { get; set; }
        public double StDevUnitWtH20 { get; set; }
        public bool BoltBlocksTogether { get; set; }
        public bool AddToeSupport { get; set; }
        public double SupportForce { get; set; }
        public string NatureOfForceApplication { get; set; }
        public double Magnitude { get; set; }
        public double Orientation { get; set; }
        public double OptimumOrientationAgainstSliding { get; set; }
        public double OptimumOrientationAgainstToppling { get; set; }
        public double EffectiveWidth { get; set; }

        // Kinematic Counters
        public int KineBlockTopCount { get; set; }  // Count of kinematically feasible block toppling
        public int KineFlexCount { get; set; }      // Count of kinematically feasible flexural toppling
        public int KineSlideCount { get; set; }     // Count of kinematically feasible sliding
        public int KineSlideTopCount { get; set; }  // Count of kinematically feasible sliding and toppling

        // Dynamic Counters
        public int SlideCount { get; set; }         // Total count of sliding mode
        public int ToppleCount { get; set; }        // Total count of toppling mode
        public int InvalidCount { get; set; }       // Count of infinite safety factors

        // Failure Type Counters
        public int StableSlideCount { get; set; }   // Count of stable sliding
        public int UnstableSlideCount { get; set; } // Count of unstable sliding
        public int StableTopCount { get; set; }     // Count of stable toppling
        public int UnStableTopCount { get; set; }   // Count of unstable toppling  

        // Constructor, initialize counters
        public GlobalParameter()
        {
            // Initialize all counters to 0
            KineBlockTopCount = 0;
            KineFlexCount = 0;
            KineSlideCount = 0;
            KineSlideTopCount = 0;
            SlideCount = 0;
            ToppleCount = 0;
            InvalidCount = 0;
            StableSlideCount = 0;
            UnstableSlideCount = 0;
            StableTopCount = 0;
            UnStableTopCount = 0;
            SupportForce = 0;
            DistSpaceA = "";
            DistSpaceB = "";
            DistFricA = "";
            DistFricB = "";
            DistSeis = "";
            DistPorePress = "";
            DistUnitWeight = "";
            DistUnitWeightH2O = "";
            NatureOfForceApplication = "Passive";
        }
    }
    public struct BlockXY
    {
        public double xTL, xBL, xBR, xTR;
        public double yTL, yBL, yBR, yTR;
    }

    public struct PorePressXY
    {
        public double xDS, yDS, xUS, yUS;
    }

    public struct MonteCarloIterationResult
    {
        public int IterationNumber;
        public double DipA;
        public double DipDirA;
        public double DipB;
        public double DipDirB;
        public double FricA;
        public double FricB;
        public bool KineBlockTop;
        public bool KineSlide;
        public bool KineFlex;
        public bool KineSlideTop;
    }
}
