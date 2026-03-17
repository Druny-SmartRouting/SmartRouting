using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using Google.OrTools.ConstraintSolver;
using SmartRouting.Types;

namespace SmartRouting.Services
{
    public class ActivityRecord
    {
        [JsonPropertyName("date")] public string Date { get; set; }[JsonPropertyName("day_of_week")] public string DayOfWeek { get; set; }[JsonPropertyName("start_time")] public string StartTime { get; set; }
        [JsonPropertyName("end_time")] public string EndTime { get; set; }
        [JsonPropertyName("technician_name")] public string TechnicianName { get; set; }
        [JsonPropertyName("location_name")] public string LocationName { get; set; }
        [JsonPropertyName("location_to")] public string LocationTo { get; set; }
        [JsonPropertyName("activity_type")] public string ActivityType { get; set; }
    }

    public class OrToolsService
    {
        public string SolveAndReturnJson(OrToolsData data)
        {
            int totalNodes = data.TimeMatrix.GetLength(0);

            RoutingIndexManager manager = new RoutingIndexManager(
                totalNodes, 
                data.VehiclesNumber, 
                data.Starts, 
                data.Ends);

            RoutingModel routing = new RoutingModel(manager);
            Solver solver = routing.solver(); 

            int transitCallbackIndex = routing.RegisterTransitCallback((long fromIndex, long toIndex) =>
            {
                var fromNode = manager.IndexToNode(fromIndex);
                var toNode = manager.IndexToNode(toIndex);
                return data.ServiceDurations[fromNode] + data.TimeMatrix[fromNode, toNode];
            });

            routing.SetArcCostEvaluatorOfAllVehicles(transitCallbackIndex);
            routing.AddDimension(transitCallbackIndex, 1440, 1440, false, "Time");
            RoutingDimension timeDimension = routing.GetMutableDimension("Time");

            for (int i = 0; i < totalNodes; ++i)
            {
                if (IsDepot(i, data.Starts, data.Ends)) continue;

                long index = manager.NodeToIndex(i);
                
                timeDimension.CumulVar(index).SetRange(0, 1440);
                routing.AddDisjunction(new long[] { index }, 100000);

                long[] openTimes = new long[data.VehiclesNumber];
                long[] closeTimes = new long[data.VehiclesNumber];

                for (int v = 0; v < data.VehiclesNumber; ++v)
                {
                    int dayOfWeek = data.TechAssignedDay[v] % 7;
                    openTimes[v] = data.SiteOpenTimes[i, dayOfWeek];
                    closeTimes[v] = data.SiteCloseTimes[i, dayOfWeek];

                    bool canServe = true;

                    ServiceType reqType = data.SiteRequiredServiceType[i];
                    SkillLevel reqLevel = data.SiteRequiredSkillLevel[i];
                    SkillLevel techLevel = SkillLevel.None;

                    if (reqType == ServiceType.Interior) techLevel = data.TechSkillInterior[v];
                    else if (reqType == ServiceType.Exterior) techLevel = data.TechSkillExterior[v];
                    else if (reqType == ServiceType.Floral) techLevel = data.TechSkillFloral[v];

                    if (techLevel < reqLevel) canServe = false;
                    if (data.SiteRequiresPhysical[i] && !data.TechCanDoPhysical[v]) canServe = false;
                    if (data.SiteRequiresLivingWalls[i] && !data.TechSkilledLivingWalls[v]) canServe = false;
                    if (data.SiteRequiresHeights[i] && !data.TechComfortableHeights[v]) canServe = false;
                    if (data.SiteRequiresLift[i] && !data.TechCertifiedLift[v]) canServe = false;
                    if (data.SiteRequiresPesticides[i] && !data.TechCertifiedPesticides[v]) canServe = false;
                    if (data.SiteRequiresCitizen[i] && !data.TechIsCitizen[v]) canServe = false;
                    
                    int techDay = data.TechAssignedDay[v];
                    if (techDay < data.SiteAllowedDays[i, 0] || techDay > data.SiteAllowedDays[i, 1]) canServe = false;
                    if (openTimes[v] == 0 && closeTimes[v] == 0) canServe = false;
                    if (!data.TechAllowedAtSite[i, v]) canServe = false;

                    if (!canServe)
                    {
                        routing.VehicleVar(index).RemoveValue(v);
                    }
                }

                IntVar isActive = routing.ActiveVar(index);
                IntVar safeVehicleVar = solver.MakeMax(routing.VehicleVar(index), 0).Var(); 

                IntExpr openExpr = solver.MakeElement(openTimes, safeVehicleVar);
                IntExpr closeExpr = solver.MakeElement(closeTimes, safeVehicleVar);

                IntExpr penalty = (1 - isActive) * 100000;

                solver.Add(timeDimension.CumulVar(index) >= openExpr - penalty);
                solver.Add(timeDimension.CumulVar(index) + data.ServiceDurations[i] <= closeExpr + penalty);
            }

            long[] nodeVisitTransit = new long[routing.Size()];
            for (int i = 0; i < routing.Size(); ++i)
            {
                int node = manager.IndexToNode(i);
                nodeVisitTransit[i] = data.ServiceDurations[node];
            }

            var techBreaks = new Dictionary<int, IntervalVar>();

            for (int i = 0; i < data.VehiclesNumber; ++i)
            {
                long startIndex = routing.Start(i);
                long endIndex = routing.End(i);

                timeDimension.CumulVar(startIndex).SetRange(data.VehicleTimeWindows[i, 0], data.VehicleTimeWindows[i, 1]);
                timeDimension.CumulVar(endIndex).SetRange(data.VehicleTimeWindows[i, 0], data.VehicleTimeWindows[i, 1]);

                IntervalVarVector breaks = new IntervalVarVector();
                IntervalVar breakInterval = solver.MakeFixedDurationIntervalVar(
                    data.BreakWindows[i, 0], 
                    data.BreakWindows[i, 1], 
                    data.Breaks[i],          
                    true,                   
                    $"Break_Tech_{i}");
                
                breaks.Add(breakInterval);
                techBreaks[i] = breakInterval;
                timeDimension.SetBreakIntervalsOfVehicle(breaks, i, nodeVisitTransit);

                IntExpr routeDuration = timeDimension.CumulVar(endIndex) - timeDimension.CumulVar(startIndex);
                solver.Add(routeDuration <= data.VehicleMaxHoursPerDay[i] * 60);
            }

            var weeklyGroups = new Dictionary<string, List<int>>();
            for (int i = 0; i < data.VehiclesNumber; i++)
            {
                string key = $"{data.VehicleToRealTech[i]}_{data.VehicleWeekNumber[i]}";
                if (!weeklyGroups.ContainsKey(key)) weeklyGroups[key] = new List<int>();
                weeklyGroups[key].Add(i);
            }

            foreach (var kvp in weeklyGroups)
            {
                string[] parts = kvp.Key.Split('_');
                int realTechId = int.Parse(parts[0]);
                int maxWeeklyMinutes = data.RealTechMaxHoursPerWeek[realTechId] * 60;

                IntExpr weeklyWork = solver.MakeIntConst(0);
                foreach (int v in kvp.Value)
                {
                    long startIndex = routing.Start(v);
                    long endIndex = routing.End(v);
                    weeklyWork = weeklyWork + (timeDimension.CumulVar(endIndex) - timeDimension.CumulVar(startIndex));
                }
                
                solver.Add(weeklyWork <= maxWeeklyMinutes);
            }

            RoutingSearchParameters searchParameters = operations_research_constraint_solver.DefaultRoutingSearchParameters();
            searchParameters.FirstSolutionStrategy = FirstSolutionStrategy.Types.Value.PathCheapestArc;
            searchParameters.TimeLimit = new Google.Protobuf.WellKnownTypes.Duration { Seconds = 15 };

            Assignment solution = routing.SolveWithParameters(searchParameters);

            if (solution != null)
            {
                return FormatSolutionJson(data, routing, manager, solution, timeDimension, techBreaks);
            }
            
            return JsonSerializer.Serialize(new { error = "Solution haven't been found. Constraints are mathematically impossible." });
        }

        private bool IsDepot(int node, int[] starts, int[] ends)
        {
            return Array.Exists(starts, s => s == node) || Array.Exists(ends, e => e == node);
        }

        private string FormatTime(long minutes)
        {
            long h = minutes / 60;
            long m = minutes % 60;
            return $"{h:D2}:{m:D2}";
        }

        private string FormatSolutionJson(OrToolsData data, RoutingModel routing, RoutingIndexManager manager, Assignment solution, RoutingDimension timeDimension, Dictionary<int, IntervalVar> techBreaks)
        {
            var allActivities = new List<ActivityRecord>();
            
            DateTime baseDate = new DateTime(2026, 3, 16); 

            for (int i = 0; i < data.VehiclesNumber; ++i)
            {
                long index = routing.Start(i);
                if (routing.IsEnd(solution.Value(routing.NextVar(index)))) continue; 

                string techName = data.TechNames[i];
                int dayOffset = data.TechAssignedDay[i];
                DateTime currentDate = baseDate.AddDays(dayOffset);
                string dateStr = currentDate.ToString("yyyy-MM-dd");
                string dayOfWeek = currentDate.DayOfWeek.ToString();

                var techActivities = new List<(long StartMin, ActivityRecord Record)>();

                while (!routing.IsEnd(index))
                {
                    long nextIndex = solution.Value(routing.NextVar(index));
                    int fromNode = manager.IndexToNode(index);
                    int toNode = manager.IndexToNode(nextIndex);

                    long departureTime = solution.Min(timeDimension.CumulVar(index)) + data.ServiceDurations[fromNode];
                    long travelTime = data.TimeMatrix[fromNode, toNode];
                    long arrivalTime = departureTime + travelTime; 

                    string fromName = routing.IsStart(index) ? "Office" : data.NodeNames[fromNode];
                    string toName = routing.IsEnd(nextIndex) ? "Office" : data.NodeNames[toNode];

                    techActivities.Add((departureTime, new ActivityRecord
                    {
                        Date = dateStr,
                        DayOfWeek = dayOfWeek,
                        StartTime = FormatTime(departureTime),
                        EndTime = FormatTime(arrivalTime),
                        TechnicianName = techName,
                        LocationName = fromName,
                        LocationTo = toName,
                        ActivityType = "Commute"
                    }));

                    if (!routing.IsEnd(nextIndex))
                    {
                        long serviceStart = solution.Min(timeDimension.CumulVar(nextIndex));
                        long serviceEnd = serviceStart + data.ServiceDurations[toNode];
                        string activityType = data.SiteRequiredServiceType[toNode].ToString();

                        techActivities.Add((serviceStart, new ActivityRecord
                        {
                            Date = dateStr,
                            DayOfWeek = dayOfWeek,
                            StartTime = FormatTime(serviceStart),
                            EndTime = FormatTime(serviceEnd),
                            TechnicianName = techName,
                            LocationName = toName,
                            LocationTo = null,
                            ActivityType = activityType
                        }));
                    }

                    index = nextIndex;
                }

                if (techBreaks.ContainsKey(i) && solution.PerformedValue(techBreaks[i]) == 1)
                {
                    long bStart = solution.StartMin(techBreaks[i]);
                    long bEnd = solution.EndMin(techBreaks[i]);
                    
                    techActivities.Add((bStart, new ActivityRecord
                    {
                        Date = dateStr,
                        DayOfWeek = dayOfWeek,
                        StartTime = FormatTime(bStart),
                        EndTime = FormatTime(bEnd),
                        TechnicianName = techName,
                        LocationName = "En Route / Break",
                        LocationTo = null,
                        ActivityType = "Break"
                    }));
                }

                var sortedTechActivities = techActivities.OrderBy(a => a.StartMin).Select(a => a.Record);
                allActivities.AddRange(sortedTechActivities);
            }

            var options = new JsonSerializerOptions { WriteIndented = true };
            return JsonSerializer.Serialize(allActivities, options);
        }
    }
}
