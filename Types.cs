namespace SmartRouting.Types
{
    public enum SkillLevel { Junior = 0, Medior = 1, Senior = 2 }

    public class Coordinates
    {
        public double Latitude { get; set; }
        public double Longitude { get; set; }
    }

    public class OrToolsData
    {
        public int[,] TimeMatrix { get; set; } = null!;
        public int[] ServiceDurations { get; set; } = null!;
        public int[,] TimeWindows { get; set; } = null!;
        public int VehiclesNumber { get; set; }
        public int[] Starts { get; set; } = null!;
        public int[] Ends { get; set; } = null!;
        public int[,] VehicleTimeWindows { get; set; } = null!;
        public int[] Breaks { get; set; } = null!;
        public int[,] BreakWindows { get; set; } = null!;

        public int[,] SiteAllowedDays { get; set; } = null!;
        public int[] TechAssignedDay { get; set; } = null!;

        public string[] TechNames { get; set; } = null!;
        public string[] NodeNames { get; set; } = null!;

        // Technician Capabilities
        public SkillLevel[] TechSkillLevel { get; set; } = null!;
        public bool[] TechCanDoPhysical { get; set; } = null!;
        public bool[] TechSkilledLivingWalls { get; set; } = null!;
        public bool[] TechComfortableHeights { get; set; } = null!;
        public bool[] TechCertifiedLift { get; set; } = null!;
        public bool[] TechCertifiedPesticides { get; set; } = null!;
        public bool[] TechIsCitizen { get; set; } = null!;

        // Site Requirements
        public SkillLevel[] SiteRequiredSkillLevel { get; set; } = null!;
        public bool[] SiteRequiresPhysical { get; set; } = null!;
        public bool[] SiteRequiresLivingWalls { get; set; } = null!;
        public bool[] SiteRequiresHeights { get; set; } = null!;
        public bool[] SiteRequiresLift { get; set; } = null!;
        public bool[] SiteRequiresPesticides { get; set; } = null!;
        public bool[] SiteRequiresCitizen { get; set; } = null!;
    }
}
