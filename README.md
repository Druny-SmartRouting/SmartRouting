# Smart Routing Web API
To run this service you will need .NET SDK 8.0 and these packages:
```
> Amazon.Lambda.AspNetCoreServer.Hosting      1.10.0
> AWSSDK.LocationService                      4.0.3.15
> ClosedXML                                   0.105.0
> Google.OrTools                              9.15.6755
```

## How to run
Just run `dotnet build` and `dotnet run`. Web server should run and you can get access to service by `localhost:5000/optimize` endpoint. You'll also need to have configurated `aws cli` in order to access Amazon Location Service.

We also have our service deployed on AWS Lambda to conveniently use it without deploying and configuring `aws cli` locally. As this repository remains private, we feel safe to leave link here: https://zxjzl3mdbffmeajjfoyxr666wu0nfxuc.lambda-url.eu-north-1.on.aws/optimize

## Input data
To get text output of a schedule you need to send POST request to said endpoint with `xlsx` file of coorect scheme. For testing we used "Routing pilot data input (Biome) -all services - FINAL.xlsx". As file parser is sensitive to sheets names and column numbers, it might not work with another file.
