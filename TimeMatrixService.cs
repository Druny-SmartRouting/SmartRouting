using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Amazon.LocationService;
using Amazon.LocationService.Model;
using SmartRouting.Types;

namespace SmartRouting.Services
{
    public class GeoTimeMatrixService
    {
        private readonly AmazonLocationServiceClient _locationClient;
        private const double EarthRadiusKm = 6371.0;
        
        private readonly string _cacheFilePath;
        private Dictionary<string, Coordinates> _cache;

        public GeoTimeMatrixService(AmazonLocationServiceClient locationClient, string cacheFilePath = "cached_coordinates.json")
        {
            _locationClient = locationClient;
            _cacheFilePath = cacheFilePath;
            _cache = new Dictionary<string, Coordinates>();
            
            LoadCache();
        }

        public async Task<long[,]> GenerateMatrixFromAddressesAsync(
            List<string> addresses, 
            double averageSpeedKmh = 30.0, 
            double routeDetourFactor = 1.3)
        {
            List<Coordinates> locations = new List<Coordinates>();

            foreach (var address in addresses)
            {
                var coords = await GetCoordinatesAsync(address);
                if (coords == null)
                {
                    throw new Exception($"Failed to get coordinates for address: {address}. Check if the address is valid.");
                }
                locations.Add(coords);
            }

            return GenerateTimeMatrix(locations, averageSpeedKmh, routeDetourFactor);
        }

        private async Task<Coordinates?> GetCoordinatesAsync(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return null;

            if (_cache.TryGetValue(address, out var cachedCoords))
            {
                Console.WriteLine($"[Cache Hit] Loaded from local file: {address}");
                return cachedCoords;
            }

            Console.WriteLine($"[Cache Miss] Calling AWS Location Service for: {address}");
            
            var request = new SearchPlaceIndexForTextRequest
            {
                IndexName = "GeoIndex", 
                Text = address,
                MaxResults = 1 
            };

            try
            {
                var response = await _locationClient.SearchPlaceIndexForTextAsync(request);

                if (response.Results.Count > 0)
                {
                    var point = response.Results[0].Place.Geometry.Point;
                    var newCoords = new Coordinates
                    {
                        Latitude = point[1], 
                        Longitude = point[0]
                    };

                    _cache[address] = newCoords;
                    SaveCache();

                    return newCoords;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"AWS Location Service rejected the request for '{address}'. REAL ERROR: {ex.Message}", ex);
            }
            return null;
        }

        private void LoadCache()
        {
            try
            {
                if (File.Exists(_cacheFilePath))
                {
                    string json = File.ReadAllText(_cacheFilePath);
                    var loadedCache = JsonSerializer.Deserialize<Dictionary<string, Coordinates>>(json);
                    if (loadedCache != null)
                    {
                        _cache = loadedCache;
                        Console.WriteLine($"[System] Loaded {_cache.Count} addresses from local cache.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[System] Warning: Could not load cache file. Starting fresh. Error: {ex.Message}");
            }
        }

        private void SaveCache()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_cache, options);
                File.WriteAllText(_cacheFilePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[System] Warning: Could not save to cache file. Error: {ex.Message}");
            }
        }

        private long[,] GenerateTimeMatrix(List<Coordinates> locations, double averageSpeedKmh, double routeDetourFactor)
        {
            int n = locations.Count;
            long[,] timeMatrix = new long[n, n];

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    if (i == j) 
                    {
                        timeMatrix[i, j] = 0;
                    }
                    else 
                    {
                        double straightDistanceKm = CalculateDistance(
                            locations[i].Latitude, locations[i].Longitude,
                            locations[j].Latitude, locations[j].Longitude);

                        double actualDistanceKm = straightDistanceKm * routeDetourFactor;
                        double timeHours = actualDistanceKm / averageSpeedKmh;
                        
                        timeMatrix[i, j] = (long)Math.Round(timeHours * 60);
                    }
                }
            }
            return timeMatrix;
        }

        private double CalculateDistance(double lat1, double lon1, double lat2, double lon2)
        {
            double dLat = DegreesToRadians(lat2 - lat1);
            double dLon = DegreesToRadians(lon2 - lon1);

            lat1 = DegreesToRadians(lat1);
            lat2 = DegreesToRadians(lat2);

            double a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
                       Math.Cos(lat1) * Math.Cos(lat2) *
                       Math.Sin(dLon / 2) * Math.Sin(dLon / 2);

            double c = 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));

            return EarthRadiusKm * c;
        }

        private double DegreesToRadians(double degrees)
        {
            return degrees * Math.PI / 180.0;
        }
    }
}
