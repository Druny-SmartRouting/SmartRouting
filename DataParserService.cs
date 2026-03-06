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

            var frequencies = rawSites.Select(s => s.FrequencyDays).Where(f => f > 0).Distinct().ToList();
            int horizonDays = CalculateLCM(frequencies);

            int totalWorkDays = horizonDays; 

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

        private Dictionary<string, int> BuildColumnMap(IXLWorksheet sheet)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int maxCol = sheet.LastCellUsed()?.Address.ColumnNumber ?? 50;

            string currentDay = null;
            string[] daysOfWeek = { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" };

            for (int col = 1; col <= maxCol; col++)
            {
                string val1 = sheet.Cell(1, col).GetString().Trim();
                string val2 = sheet.Cell(2, col).GetString().Trim();
                string val3 = sheet.Cell(3, col).GetString().Trim();

                var foundDay = daysOfWeek.FirstOrDefault(d => val2.StartsWith(d, StringComparison.OrdinalIgnoreCase));
                if (foundDay != null)
                {
                    currentDay = foundDay;
                }

                if (currentDay != null && val3.Equals("from", StringComparison.OrdinalIgnoreCase))
                {
                    map[$"{currentDay}_from"] = col;
                }
                else if (currentDay != null && val3.Equals("to", StringComparison.OrdinalIgnoreCase))
                {
                    map[$"{currentDay}_to"] = col;
                }
                else
                {
                    if (!string.IsNullOrEmpty(val2) && foundDay == null)
                    {
                        map[NormalizeHeader(val2)] = col;
                    }
                    else if (!string.IsNullOrEmpty(val1) && string.IsNullOrEmpty(val2))
                    {
                        map[NormalizeHeader(val1)] = col;
                    }
                }
            }
            return map;
        }

        private string NormalizeHeader(string header)
        {
            if (string.IsNullOrWhiteSpace(header)) return string.Empty;
            return Regex.Replace(header.ToLowerInvariant(), @"[^a-z0-9]", "");
        }

        private int GetColIndex(Dictionary<string, int> map, params string[] possibleNames)
        {
            foreach (var name in possibleNames)
            {
                var normalizedName = NormalizeHeader(name);
                
                if (map.TryGetValue(normalizedName, out int col))
                    return col;

                var partialMatch = map.Keys.FirstOrDefault(k => k.Contains(normalizedName) || normalizedName.Contains(k));
                if (partialMatch != null)
                    return map[partialMatch];
            }
            return -1;
        }

        private string GetString(IXLWorksheet sheet, int row, int col)
        {
            if (col <= 0) return string.Empty;
            return sheet.Cell(row, col).GetString()?.Trim() ?? string.Empty;
        }

        private List<RawSite> ReadRawSites(IXLWorksheet sheet)
        {
            var sites = new List<RawSite>();
            var map = BuildColumnMap(sheet);

            int colName = GetColIndex(map, "site name or code", "name");
            int colAddress = GetColIndex(map, "site address");
            int colFreq = GetColIndex(map, "visit frequency");
            int colDuration = GetColIndex(map, "est duration");
            int colSkill = GetColIndex(map, "service skill requirement");
            int colPhys = GetColIndex(map, "physically demanding");
            int colWalls = GetColIndex(map, "living walls");
            int colHeights = GetColIndex(map, "work at heights");
            int colLift = GetColIndex(map, "using the lift");
            int colPest = GetColIndex(map, "application of pesticides");
            int colCit = GetColIndex(map, "citizen technician");

            string[] daysOfWeek = { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" };

            int row = 4;
            while (!sheet.Cell(row, colName > 0 ? colName : 1).IsEmpty())
            {
                var site = new RawSite
                {
                    Name = GetString(sheet, row, colName),
                    Address = GetString(sheet, row, colAddress),
                    FrequencyDays = ParseFrequency(GetString(sheet, row, colFreq)),
                    Duration = int.TryParse(GetString(sheet, row, colDuration), out var d) ? d : 30,
                    SkillStr = GetString(sheet, row, colSkill),
                    ReqPhysical = ParseBool(GetString(sheet, row, colPhys)),
                    ReqLivingWalls = ParseBool(GetString(sheet, row, colWalls)),
                    ReqHeights = ParseBool(GetString(sheet, row, colHeights)),
                    ReqLift = ParseBool(GetString(sheet, row, colLift)),
                    ReqPesticides = ParseBool(GetString(sheet, row, colPest)),
                    ReqCitizen = ParseBool(GetString(sheet, row, colCit))
                };

                for (int day = 0; day < 7; day++)
                {
                    map.TryGetValue($"{daysOfWeek[day]}_from", out int colOpen);
                    map.TryGetValue($"{daysOfWeek[day]}_to", out int colClose);

                    site.OpenTimes[day] = ParseTime(GetString(sheet, row, colOpen), -1);
                    site.CloseTimes[day] = ParseTime(GetString(sheet, row, colClose), -1);
                    
                    if (site.OpenTimes[day] == -1 || site.CloseTimes[day] == -1) {
                        site.OpenTimes[day] = 0;
                        site.CloseTimes[day] = 0;
                    }
                }
                sites.Add(site);
                row++;
            }
            return sites;
        }

        private List<RawTech> ReadRawTechs(IXLWorksheet sheet)
        {
            var techs = new List<RawTech>();
            var map = BuildColumnMap(sheet);

            int colName = GetColIndex(map, "name");
            int colHome = GetColIndex(map, "home address");
            int colOffice = GetColIndex(map, "office address");
            int colStartFrom = GetColIndex(map, "starts from");
            int colFinishAt = GetColIndex(map, "finishes at");
            int colBreakDur = GetColIndex(map, "min break");
            int colBreakStart = GetColIndex(map, "not earlier than", "break start");
            int colBreakEnd = GetColIndex(map, "not later than", "break end");
            int colSkill = GetColIndex(map, "service skills");
            int colPhys = GetColIndex(map, "physically demanding");
            int colWalls = GetColIndex(map, "living walls");
            int colHeights = GetColIndex(map, "work at heights");
            int colLift = GetColIndex(map, "using the lift");
            int colPest = GetColIndex(map, "pesticide applicator");
            int colCit = GetColIndex(map, "is a citizen");

            string[] daysOfWeek = { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" };

            int row = 4;
            while (!sheet.Cell(row, colName > 0 ? colName : 1).IsEmpty())
            {
                string home = GetString(sheet, row, colHome);
                string office = GetString(sheet, row, colOffice);
                string startLocStr = GetString(sheet, row, colStartFrom).ToLower();
                string endLocStr = GetString(sheet, row, colFinishAt).ToLower();

                string startLoc = startLocStr.Contains("office") ? office : home;
                string endLoc = endLocStr.Contains("office") ? office : home;

                var tech = new RawTech
                {
                    Name = GetString(sheet, row, colName),
                    StartAddress = startLoc,
                    EndAddress = endLoc,
                    
                    BreakDuration = int.TryParse(GetString(sheet, row, colBreakDur), out var b) ? b : 30,
                    BreakStart = ParseTime(GetString(sheet, row, colBreakStart), 720),
                    BreakEnd = ParseTime(GetString(sheet, row, colBreakEnd), 840),
                    SkillsStr = GetString(sheet, row, colSkill),
                    CanPhysical = ParseBool(GetString(sheet, row, colPhys)),
                    CanLivingWalls = ParseBool(GetString(sheet, row, colWalls)),
                    CanHeights = ParseBool(GetString(sheet, row, colHeights)),
                    CanLift = ParseBool(GetString(sheet, row, colLift)),
                    CanPesticides = ParseBool(GetString(sheet, row, colPest)),
                    IsCitizen = ParseBool(GetString(sheet, row, colCit))
                };

                for (int day = 0; day < 7; day++)
                {
                    map.TryGetValue($"{daysOfWeek[day]}_from", out int colOpen);
                    map.TryGetValue($"{daysOfWeek[day]}_to", out int colClose);

                    tech.ShiftStarts[day] = ParseTime(GetString(sheet, row, colOpen), -1);
                    tech.ShiftEnds[day] = ParseTime(GetString(sheet, row, colClose), -1);
                    
                    if (tech.ShiftStarts[day] == -1 || tech.ShiftEnds[day] == -1) {
                        tech.ShiftStarts[day] = 0;
                        tech.ShiftEnds[day] = 0;
                    }
                }
                techs.Add(tech);
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
                    int dayOfWeek = day % 7;
                    if (tech.ShiftStarts[dayOfWeek] == 0 && tech.ShiftEnds[dayOfWeek] == 0) continue;

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
                
                TechSkillInterior = new SkillLevel[numVehicles],
                TechSkillExterior = new SkillLevel[numVehicles],
                TechSkillFloral = new SkillLevel[numVehicles],
                
                TechCanDoPhysical = new bool[numVehicles],
                TechSkilledLivingWalls = new bool[numVehicles],
                TechComfortableHeights = new bool[numVehicles],
                TechCertifiedLift = new bool[numVehicles],
                TechCertifiedPesticides = new bool[numVehicles],
                TechIsCitizen = new bool[numVehicles],
                
                SiteRequiredServiceType = new ServiceType[numNodes],
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
                NodeNames = new string[numNodes],
                
                SiteOpenTimes = new int[numNodes, 7],
                SiteCloseTimes = new int[numNodes, 7]
            };

            data.TimeMatrix = new int[numNodes, numNodes];
            List<string> nodeAddresses = new List<string>();
            
            for (int i = 0; i < numVehicles; i++) {
                nodeAddresses.Add(vTechs[i].Tech.StartAddress);
                data.Starts[i] = i; 
                data.TechNames[i] = vTechs[i].Tech.Name; 
                data.NodeNames[i] = $"{vTechs[i].Tech.Name}'s Start Location (Day {vTechs[i].AssignedDay})"; 
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
                int dayOfWeek = vt.AssignedDay % 7;

                data.VehicleTimeWindows[i, 0] = vt.Tech.ShiftStarts[dayOfWeek];
                data.VehicleTimeWindows[i, 1] = vt.Tech.ShiftEnds[dayOfWeek];
                data.Breaks[i] = vt.Tech.BreakDuration;
                data.BreakWindows[i, 0] = vt.Tech.BreakStart;
                data.BreakWindows[i, 1] = vt.Tech.BreakEnd;
                data.TechAssignedDay[i] = vt.AssignedDay;

                ParseTechSkills(vt.Tech.SkillsStr, out var iLvl, out var eLvl, out var fLvl);
                data.TechSkillInterior[i] = iLvl;
                data.TechSkillExterior[i] = eLvl;
                data.TechSkillFloral[i] = fLvl;
                
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
                    
                    for (int day = 0; day < 7; day++) {
                        data.SiteOpenTimes[i, day] = 0;
                        data.SiteCloseTimes[i, day] = 1440;
                    }
                }
                else
                {
                    var vs = vSites[i - (numVehicles * 2)];
                    data.ServiceDurations[i] = vs.Site.Duration;
                    
                    data.SiteAllowedDays[i, 0] = vs.AllowedStartDay;
                    data.SiteAllowedDays[i, 1] = vs.AllowedEndDay;

                    ParseSiteSkill(vs.Site.SkillStr, out var sType, out var sLvl);
                    data.SiteRequiredServiceType[i] = sType;
                    data.SiteRequiredSkillLevel[i] = sLvl;
                    
                    data.SiteRequiresPhysical[i] = vs.Site.ReqPhysical;
                    data.SiteRequiresLivingWalls[i] = vs.Site.ReqLivingWalls;
                    data.SiteRequiresHeights[i] = vs.Site.ReqHeights;
                    data.SiteRequiresLift[i] = vs.Site.ReqLift;
                    data.SiteRequiresPesticides[i] = vs.Site.ReqPesticides;
                    data.SiteRequiresCitizen[i] = vs.Site.ReqCitizen;

                    for (int day = 0; day < 7; day++) {
                        data.SiteOpenTimes[i, day] = vs.Site.OpenTimes[day];
                        data.SiteCloseTimes[i, day] = vs.Site.CloseTimes[day];
                    }
                }
            }

            return data;
        }

        private int CalculateLCM(List<int> numbers)
        {
            if (numbers == null || numbers.Count == 0) return 14;
            return numbers.Aggregate(Lcm);
        }

        private int Lcm(int a, int b)
        {
            if (a == 0 || b == 0) return 0;
            return (a / Gcd(a, b)) * b;
        }

        private int Gcd(int a, int b)
        {
            while (b != 0)
            {
                int temp = b;
                b = a % b;
                a = temp;
            }
            return a;
        }

        private int ParseFrequency(string freq)
        {
            var match = Regex.Match(freq ?? "", @"\d+(?=\s*days)");
            return match.Success ? int.Parse(match.Value) : 14; 
        }

        private void ParseTechSkills(string skillsStr, out SkillLevel interior, out SkillLevel exterior, out SkillLevel floral)
        {
            interior = SkillLevel.None; exterior = SkillLevel.None; floral = SkillLevel.None;
            if (string.IsNullOrWhiteSpace(skillsStr)) return;
            
            var parts = skillsStr.ToLower().Split(',');
            foreach (var p in parts)
            {
                if (p.Contains("interior")) interior = ExtractLevel(p);
                if (p.Contains("exterior")) exterior = ExtractLevel(p);
                if (p.Contains("floral")) floral = ExtractLevel(p);
            }
        }

        private void ParseSiteSkill(string skillStr, out ServiceType type, out SkillLevel level)
        {
            type = ServiceType.Interior; 
            level = SkillLevel.Junior; 
            if (string.IsNullOrWhiteSpace(skillStr)) return;
            
            var s = skillStr.ToLower();
            if (s.Contains("exterior")) type = ServiceType.Exterior;
            else if (s.Contains("floral")) type = ServiceType.Floral;
            
            level = ExtractLevel(s);
        }

        private SkillLevel ExtractLevel(string s)
        {
            if (s.Contains("senior")) return SkillLevel.Senior;
            if (s.Contains("medior")) return SkillLevel.Medior;
            if (s.Contains("junior")) return SkillLevel.Junior;
            return SkillLevel.None;
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
        public int[] OpenTimes { get; set; } = new int[7];  
        public int[] CloseTimes { get; set; } = new int[7]; 
        public int FrequencyDays { get; set; }
        public int Duration { get; set; }
        public string SkillStr { get; set; }
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
        public int[] ShiftStarts { get; set; } = new int[7]; 
        public int[] ShiftEnds { get; set; } = new int[7];   
        public int BreakDuration { get; set; }
        public int BreakStart { get; set; }
        public int BreakEnd { get; set; }
        public string SkillsStr { get; set; }
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
