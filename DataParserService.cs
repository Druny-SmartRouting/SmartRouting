using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;
using SmartRouting.Types;

namespace SmartRouting.Services
{
    public class DataParserService
    {
        private readonly GeoTimeMatrixService _geoService;

        public DataParserService(GeoTimeMatrixService geoService)
        {
            _geoService = geoService;
        }

        public async Task<OrToolsData> ParseExcelAndGenerateDataAsync(System.IO.Stream fileStream)
        {
            using var workbook = new XLWorkbook(fileStream);
            var sitesSheet = workbook.Worksheet("Service Sites");
            var techsSheet = workbook.Worksheet("Technicians");

            var rawSites = ReadRawSites(sitesSheet);
            var rawTechs = ReadRawTechs(techsSheet);

            int horizonDays = rawSites.Max(s => s.FrequencyDays);
            int totalWorkDays = (horizonDays / 7) * 5; 

            var virtualTechs = GenerateVirtualTechs(rawTechs, totalWorkDays);
            var virtualSites = GenerateVirtualSites(rawSites, horizonDays, totalWorkDays);

            var allAddresses = new List<string>();
            
            allAddresses.AddRange(virtualTechs.Select(t => t.Tech.StartAddress));
            allAddresses.AddRange(virtualTechs.Select(t => t.Tech.EndAddress));
            allAddresses.AddRange(virtualSites.Select(s => s.Site.Address));
            
            var uniqueAddresses = allAddresses.Distinct().ToList();
            var uniqueTimeMatrix = await _geoService.GenerateMatrixFromAddressesAsync(uniqueAddresses);

            return BuildOrToolsData(virtualTechs, virtualSites, uniqueAddresses, uniqueTimeMatrix);
        }

        private List<RawSite> ReadRawSites(IXLWorksheet sheet)
        {
            var sites = new List<RawSite>();
            int row = 4;
            while (!sheet.Cell(row, 1).IsEmpty())
            {
                sites.Add(new RawSite
                {
                    Name = sheet.Cell(row, 1).GetString(),
                    Address = sheet.Cell(row, 3).GetString(),
                    OpenTime = ParseTime(sheet.Cell(row, 10).GetString(), 480),
                    CloseTime = ParseTime(sheet.Cell(row, 11).GetString(), 1020),
                    FrequencyDays = ParseFrequency(sheet.Cell(row, 24).GetString()),
                    Duration = sheet.Cell(row, 25).TryGetValue<int>(out var d) ? d : 30, 
                    Skill = ParseSkillLevel(sheet.Cell(row, 26).GetString()),
                    ReqPhysical = ParseBool(sheet.Cell(row, 27).GetString()),
                    ReqLivingWalls = ParseBool(sheet.Cell(row, 28).GetString()),
                    ReqHeights = ParseBool(sheet.Cell(row, 29).GetString()),
                    ReqLift = ParseBool(sheet.Cell(row, 30).GetString()),
                    ReqPesticides = ParseBool(sheet.Cell(row, 31).GetString()),
                    ReqCitizen = ParseBool(sheet.Cell(row, 32).GetString())
                });
                row++;
            }
            return sites;
        }

        private List<RawTech> ReadRawTechs(IXLWorksheet sheet)
        {
            var techs = new List<RawTech>();
            int row = 4;
            while (!sheet.Cell(row, 1).IsEmpty())
            {
                string home = sheet.Cell(row, 2).GetString();
                string office = sheet.Cell(row, 3).GetString();
                string startLoc = sheet.Cell(row, 4).GetString().ToLower().Contains("office") ? office : home;
                string endLoc = sheet.Cell(row, 5).GetString().ToLower().Contains("office") ? office : home;

                techs.Add(new RawTech
                {
                    Name = sheet.Cell(row, 1).GetString(),
                    StartAddress = startLoc,
                    EndAddress = endLoc,
                    ShiftStart = ParseTime(sheet.Cell(row, 6).GetString(), 480),
                    ShiftEnd = ParseTime(sheet.Cell(row, 7).GetString(), 1020),
                    BreakDuration = sheet.Cell(row, 20).TryGetValue<int>(out var b) ? b : 30,
                    BreakStart = ParseTime(sheet.Cell(row, 21).GetString(), 720),
                    BreakEnd = ParseTime(sheet.Cell(row, 22).GetString(), 840),
                    Skill = ParseSkillLevel(sheet.Cell(row, 25).GetString()),
                    CanPhysical = ParseBool(sheet.Cell(row, 26).GetString()),
                    CanLivingWalls = ParseBool(sheet.Cell(row, 27).GetString()),
                    CanHeights = ParseBool(sheet.Cell(row, 28).GetString()),
                    CanLift = ParseBool(sheet.Cell(row, 29).GetString()),
                    CanPesticides = ParseBool(sheet.Cell(row, 30).GetString()),
                    IsCitizen = ParseBool(sheet.Cell(row, 31).GetString())
                });
                row++;
            }
            return techs;
        }

        

        private List<VirtualTech> GenerateVirtualTechs(List<RawTech> rawTechs, int totalWorkDays)
        {
            var virtualTechs = new List<VirtualTech>();
            foreach (var tech in rawTechs)
            {
                for (int day = 0; day < totalWorkDays; day++)
                {
                    virtualTechs.Add(new VirtualTech { Tech = tech, AssignedDay = day });
                }
            }
            return virtualTechs;
        }

        private List<VirtualSite> GenerateVirtualSites(List<RawSite> rawSites, int horizonDays, int totalWorkDays)
        {
            var virtualSites = new List<VirtualSite>();
            foreach (var site in rawSites)
            {
                int visitsNeeded = horizonDays / (site.FrequencyDays == 0 ? 14 : site.FrequencyDays); 
                if (visitsNeeded == 0) visitsNeeded = 1;

                int workDaysPerVisitWindow = totalWorkDays / visitsNeeded;

                for (int visit = 0; visit < visitsNeeded; visit++)
                {
                    virtualSites.Add(new VirtualSite
                    {
                        Site = site,
                        AllowedStartDay = visit * workDaysPerVisitWindow,
                        AllowedEndDay = (visit + 1) * workDaysPerVisitWindow - 1
                    });
                }
            }
            return virtualSites;
        }

        private OrToolsData BuildOrToolsData(
            List<VirtualTech> vTechs, List<VirtualSite> vSites, 
            List<string> uniqueAddresses, long[,] uniqueMatrix)
        {
            int numVehicles = vTechs.Count;
            int numNodes = numVehicles * 2 + vSites.Count; 

            var data = new OrToolsData
            {
                VehiclesNumber = numVehicles,
                Starts = new int[numVehicles],
                Ends = new int[numVehicles],
                ServiceDurations = new int[numNodes],
                TimeWindows = new int[numNodes, 2],
                VehicleTimeWindows = new int[numVehicles, 2],
                Breaks = new int[numVehicles],
                BreakWindows = new int[numVehicles, 2],
                
                TechSkillLevel = new SkillLevel[numVehicles],
                TechCanDoPhysical = new bool[numVehicles],
                TechSkilledLivingWalls = new bool[numVehicles],
                TechComfortableHeights = new bool[numVehicles],
                TechCertifiedLift = new bool[numVehicles],
                TechCertifiedPesticides = new bool[numVehicles],
                TechIsCitizen = new bool[numVehicles],

                SiteRequiredSkillLevel = new SkillLevel[numNodes],
                SiteRequiresPhysical = new bool[numNodes],
                SiteRequiresLivingWalls = new bool[numNodes],
                SiteRequiresHeights = new bool[numNodes],
                SiteRequiresLift = new bool[numNodes],
                SiteRequiresPesticides = new bool[numNodes],
                SiteRequiresCitizen = new bool[numNodes],
                
                SiteAllowedDays = new int[numNodes, 2], 
                TechAssignedDay = new int[numVehicles],

                TechNames = new string[numVehicles],
                NodeNames = new string[numNodes]
            };

            data.TimeMatrix = new int[numNodes, numNodes];
            List<string> nodeAddresses = new List<string>();
            
            for (int i = 0; i < numVehicles; i++) {
                nodeAddresses.Add(vTechs[i].Tech.StartAddress);
                data.Starts[i] = i;
                data.TechNames[i] = vTechs[i].Tech.Name;
                data.NodeNames[i] = $"{vTechs[i].Tech.Name}'s Start Location";
            }
            for (int i = 0; i < numVehicles; i++) {
                nodeAddresses.Add(vTechs[i].Tech.EndAddress);
                data.Ends[i] = numVehicles + i;
                data.NodeNames[numVehicles + i] = $"{vTechs[i].Tech.Name}'s End Location";
            }
            for (int i = 0; i < vSites.Count; i++) {
                nodeAddresses.Add(vSites[i].Site.Address);
                data.NodeNames[numVehicles * 2 + i] = vSites[i].Site.Name;
            }

            for (int i = 0; i < numNodes; i++) {
                for (int j = 0; j < numNodes; j++) {
                    int idxI = uniqueAddresses.IndexOf(nodeAddresses[i]);
                    int idxJ = uniqueAddresses.IndexOf(nodeAddresses[j]);
                    data.TimeMatrix[i, j] = (int)uniqueMatrix[idxI, idxJ];
                }
            }

            for (int i = 0; i < numVehicles; i++)
            {
                var vt = vTechs[i];
                data.VehicleTimeWindows[i, 0] = vt.Tech.ShiftStart;
                data.VehicleTimeWindows[i, 1] = vt.Tech.ShiftEnd;
                data.Breaks[i] = vt.Tech.BreakDuration;
                data.BreakWindows[i, 0] = vt.Tech.BreakStart;
                data.BreakWindows[i, 1] = vt.Tech.BreakEnd;
                data.TechAssignedDay[i] = vt.AssignedDay;

                data.TechSkillLevel[i] = vt.Tech.Skill;
                data.TechCanDoPhysical[i] = vt.Tech.CanPhysical;
                data.TechSkilledLivingWalls[i] = vt.Tech.CanLivingWalls;
                data.TechComfortableHeights[i] = vt.Tech.CanHeights;
                data.TechCertifiedLift[i] = vt.Tech.CanLift;
                data.TechCertifiedPesticides[i] = vt.Tech.CanPesticides;
                data.TechIsCitizen[i] = vt.Tech.IsCitizen;
            }

            for (int i = 0; i < numNodes; i++)
            {
                if (i < numVehicles * 2) 
                {
                    data.ServiceDurations[i] = 0;
                    data.TimeWindows[i, 0] = 0;
                    data.TimeWindows[i, 1] = 1440;
                }
                else
                {
                    var vs = vSites[i - (numVehicles * 2)];
                    data.ServiceDurations[i] = vs.Site.Duration;
                    data.TimeWindows[i, 0] = vs.Site.OpenTime;
                    data.TimeWindows[i, 1] = vs.Site.CloseTime;
                    
                    data.SiteAllowedDays[i, 0] = vs.AllowedStartDay;
                    data.SiteAllowedDays[i, 1] = vs.AllowedEndDay;

                    data.SiteRequiredSkillLevel[i] = vs.Site.Skill;
                    data.SiteRequiresPhysical[i] = vs.Site.ReqPhysical;
                    data.SiteRequiresLivingWalls[i] = vs.Site.ReqLivingWalls;
                    data.SiteRequiresHeights[i] = vs.Site.ReqHeights;
                    data.SiteRequiresLift[i] = vs.Site.ReqLift;
                    data.SiteRequiresPesticides[i] = vs.Site.ReqPesticides;
                    data.SiteRequiresCitizen[i] = vs.Site.ReqCitizen;
                }
            }

            return data;
        }

        private int ParseFrequency(string freq)
        {
            var match = Regex.Match(freq ?? "", @"\d+(?=\s*days)");
            return match.Success ? int.Parse(match.Value) : 14; 
        }

        private SkillLevel ParseSkillLevel(string skillStr)
        {
            if (string.IsNullOrEmpty(skillStr)) return SkillLevel.Junior;
            skillStr = skillStr.ToLower();
            if (skillStr.Contains("senior")) return SkillLevel.Senior;
            if (skillStr.Contains("medior")) return SkillLevel.Medior;
            return SkillLevel.Junior;
        }

        private bool ParseBool(string val) => val?.Trim().ToLower() == "yes";

        private int ParseTime(string timeStr, int defaultMinutes)
        {
            if (string.IsNullOrWhiteSpace(timeStr)) return defaultMinutes;            
            if (DateTime.TryParse(timeStr, out DateTime dt)) return (int)dt.TimeOfDay.TotalMinutes;
            if (TimeSpan.TryParse(timeStr, out TimeSpan ts)) return (int)ts.TotalMinutes;
            
            return defaultMinutes;
        }
    }
    
    public class RawSite {
        public string Name { get; set; }
        public string Address { get; set; }
        public int OpenTime { get; set; }
        public int CloseTime { get; set; }
        public int FrequencyDays { get; set; }
        public int Duration { get; set; }
        public SkillLevel Skill { get; set; }
        public bool ReqPhysical { get; set; }
        public bool ReqLivingWalls { get; set; }
        public bool ReqHeights { get; set; }
        public bool ReqLift { get; set; }
        public bool ReqPesticides { get; set; }
        public bool ReqCitizen { get; set; }
    }

    public class RawTech {
        public string Name { get; set; }
        public string StartAddress { get; set; }
        public string EndAddress { get; set; }
        public int ShiftStart { get; set; }
        public int ShiftEnd { get; set; }
        public int BreakDuration { get; set; }
        public int BreakStart { get; set; }
        public int BreakEnd { get; set; }
        public SkillLevel Skill { get; set; }
        public bool CanPhysical { get; set; }
        public bool CanLivingWalls { get; set; }
        public bool CanHeights { get; set; }
        public bool CanLift { get; set; }
        public bool CanPesticides { get; set; }
        public bool IsCitizen { get; set; }
    }

    public class VirtualTech {
        public RawTech Tech { get; set; }
        public int AssignedDay { get; set; } 
    }

    public class VirtualSite {
        public RawSite Site { get; set; }
        public int AllowedStartDay { get; set; } 
        public int AllowedEndDay { get; set; }
    }
}
