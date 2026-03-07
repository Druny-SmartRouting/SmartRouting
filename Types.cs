namespace SmartRouting.Types
{
    public enum SkillLevel { None = 0, Junior = 1, Medior = 2, Senior = 3 }
    public enum ServiceType { Interior, Exterior, Floral }

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

        public int[,] SiteOpenTimes { get; set; } = null!;
        public int[,] SiteCloseTimes { get; set; } = null!;

        public bool[,] TechAllowedAtSite { get; set; } = null!;

        public int[] VehicleMaxHoursPerDay { get; set; } = null!; 

        public int[] VehicleToRealTech { get; set; } = null!; 
        public int[] VehicleWeekNumber { get; set; } = null!; 
        public int[] RealTechMaxHoursPerWeek { get; set; } = null!; 

        public SkillLevel[] TechSkillInterior { get; set; } = null!;
        public SkillLevel[] TechSkillExterior { get; set; } = null!;
        public SkillLevel[] TechSkillFloral { get; set; } = null!;
        
        public bool[] TechCanDoPhysical { get; set; } = null!;
        public bool[] TechSkilledLivingWalls { get; set; } = null!;
        public bool[] TechComfortableHeights { get; set; } = null!;
        public bool[] TechCertifiedLift { get; set; } = null!;
        public bool[] TechCertifiedPesticides { get; set; } = null!;
        public bool[] TechIsCitizen { get; set; } = null!;
        
        public ServiceType[] SiteRequiredServiceType { get; set; } = null!;
        public SkillLevel[] SiteRequiredSkillLevel { get; set; } = null!;
        public bool[] SiteRequiresPhysical { get; set; } = null!;
        public bool[] SiteRequiresLivingWalls { get; set; } = null!;
        public bool[] SiteRequiresHeights { get; set; } = null!;
        public bool[] SiteRequiresLift { get; set; } = null!;
        public bool[] SiteRequiresPesticides { get; set; } = null!;
        public bool[] SiteRequiresCitizen { get; set; } = null!;
    }
}
