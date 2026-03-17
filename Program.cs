using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Amazon.LocationService;
using SmartRouting.Services;
using System.IO;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddAWSLambdaHosting(LambdaEventSource.HttpApi);

builder.Services.AddSingleton<AmazonLocationServiceClient>();
builder.Services.AddSingleton(sp => 
    new GeoTimeMatrixService(sp.GetRequiredService<AmazonLocationServiceClient>(), "/tmp/cached_coordinates.json"));
builder.Services.AddSingleton<DataParserService>();
builder.Services.AddSingleton<OrToolsService>();

var app = builder.Build();

app.MapPost("/optimize", async (IFormFile file, DataParserService dataParser, OrToolsService orTools) =>
{
    if (file == null || file.Length == 0)
        return Results.BadRequest("Please upload a valid Excel file.");

    try
    {
        using var stream = new MemoryStream();
        await file.CopyToAsync(stream);
        stream.Position = 0;

        var data = await dataParser.ParseExcelAndGenerateDataAsync(stream);

        string resultJson = orTools.SolveAndReturnJson(data);

        return Results.Content(resultJson, "application/json");
    }
    catch (System.Exception ex)
    {
        if (ex.Message.Contains("Failed to get coordinates"))
        {
            return Results.BadRequest(ex.Message);
        }
        
        return Results.Problem($"An error occurred: {ex.Message}\n{ex.StackTrace}");
    }
})
.DisableAntiforgery();

app.Run();
