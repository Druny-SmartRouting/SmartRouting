using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.OrTools.ConstraintSolver;
using SmartRouting.Types;

namespace SmartRouting.Services
{
    public class OrToolsService
    {
        public string SolveAndReturnText(OrToolsData data)
        {
            StringBuilder sb = new StringBuilder();
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
                timeDimension.SetBreakIntervalsOfVehicle(breaks, i, nodeVisitTransit);
            }

            RoutingSearchParameters searchParameters = operations_research_constraint_solver.DefaultRoutingSearchParameters();
            searchParameters.FirstSolutionStrategy = FirstSolutionStrategy.Types.Value.PathCheapestArc;
            searchParameters.TimeLimit = new Google.Protobuf.WellKnownTypes.Duration { Seconds = 15 };

            Assignment solution = routing.SolveWithParameters(searchParameters);

            if (solution != null)
            {
                FormatSolution(data, routing, manager, solution, timeDimension, sb);
            }
            else
            {
                sb.AppendLine("Solution haven't been found. Constraints are mathematically impossible.");
            }

            return sb.ToString();
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

        private void FormatSolution(OrToolsData data, RoutingModel routing, RoutingIndexManager manager, Assignment solution, RoutingDimension timeDimension, StringBuilder sb)
        {
            List<int> droppedNodes = new List<int>();
            for (int index = 0; index < routing.Size(); ++index)
            {
                if (routing.IsStart(index) || routing.IsEnd(index)) continue;
                if (solution.Value(routing.NextVar(index)) == index)
                {
                    droppedNodes.Add(manager.IndexToNode(index));
                }
            }

            if (droppedNodes.Count > 0)
            {
                sb.AppendLine("WARNING: The following sites were DROPPED (Constraints too tight):");
                foreach (var dropped in droppedNodes)
                {
                    sb.AppendLine($"- {data.NodeNames[dropped]}");
                }
                sb.AppendLine();
            }

            sb.AppendLine($"Optimization complete. Full cost (drive + work + penalties): {solution.ObjectiveValue()}\n");

            var routes = new List<RouteResult>();

            for (int i = 0; i < data.VehiclesNumber; ++i)
            {
                long index = routing.Start(i);
                
                if (routing.IsEnd(solution.Value(routing.NextVar(index)))) 
                {
                    continue; 
                }

                string techName = data.TechNames[i];
                int day = data.TechAssignedDay[i];
                string routeStr = "";

                while (!routing.IsEnd(index))
                {
                    var timeVar = timeDimension.CumulVar(index);
                    long minTime = solution.Min(timeVar);
                    int node = manager.IndexToNode(index);
                    
                    if (routing.IsStart(index)) {
                        routeStr += $"[Start] {data.NodeNames[node]} (Dep: {FormatTime(minTime)}) -> ";
                    } else {
                        routeStr += $"[Visit] {data.NodeNames[node]} (Arr: {FormatTime(minTime)}, Job: {data.ServiceDurations[node]}m) -> ";
                    }
                    
                    index = solution.Value(routing.NextVar(index));
                }

                var endTimeVar = timeDimension.CumulVar(index);
                long finishTime = solution.Min(endTimeVar);
                int endNode = manager.IndexToNode(index);
                routeStr += $"[Finish] {data.NodeNames[endNode]} (Arr: {FormatTime(finishTime)})";

                routes.Add(new RouteResult 
                { 
                    TechName = techName, 
                    Day = day, 
                    RouteString = routeStr 
                });
            }

            var sortedRoutes = routes.OrderBy(r => r.Day).ThenBy(r => r.TechName).ToList();

            foreach (var r in sortedRoutes)
            {
                sb.AppendLine($"--- Day {r.Day} | Technician: {r.TechName} ---");
                sb.AppendLine(r.RouteString + "\n");
            }
        }

        private class RouteResult
        {
            public string TechName { get; set; }
            public int Day { get; set; }
            public string RouteString { get; set; }
        }
    }
}
