# Smart Routing Web API

To run this service, you will need the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) and the following packages:

```text
Amazon.Lambda.AspNetCoreServer.Hosting      1.10.0
AWSSDK.LocationService                      4.0.3.15
ClosedXML                                   0.105.0
Google.OrTools                              9.15.6755
```

## How to Run

Just run `dotnet build` and `dotnet run`. The web server should start, and you can access the service via the `http://localhost:5000/optimize` endpoint. You will also need to have the `aws cli` configured locally in order to access the Amazon Location Service.

We have also deployed our service to AWS Lambda to conveniently test it without local AWS CLI configuration, confirming it works perfectly in such an environment. AWS Lambda or other serverless options are the best choice for hosting services like this, as they run only on-demand instead of requiring a 24/7 server deployment, minimizing hosting costs for the business. We highly recommend this option for production deployment.

## Input Data

To get the text output of a schedule, you need to send a `POST` request to the endpoint with an `.xlsx` file matching the **correct schema**. For testing, we used `"Routing pilot data input (Biome) - all services - FINAL.xlsx"` and `"Routing pilot data input (Oasis) - exterior only - FINAL.xlsx"`. 

Although the data parser is flexible to some extent, you need at least 2 separate sheets named exactly `"Service Sites"` and `"Technicians"` in your file for the service to correctly identify the stored data.

## Detailed Description of the Project

During the development of our routing optimization service, we faced a complex logistical challenge: transforming the chaos of hundreds of locations, varying operating hours, specific technician skills, and overlapping visit frequencies into a clear, actionable schedule. To ensure the solution is fast, reliable, and scalable, we built the architecture around the following core technical decisions:

1. **Mathematical Core: Google OR-Tools**  
   Google OR-Tools is the industry standard for solving complex Vehicle Routing Problems (VRP). It natively supports our tech stack (C# / .NET), excels at managing strict time windows, and allows for flexible constraint configuration (e.g., matching a technician's specific certification with a site's requirements). This guarantees mathematically proven efficiency and optimal route generation.

2. **Multi-Week Schedule Modeling: The "Virtual Entities" Approach**  
   We implemented a Least Common Multiple (LCM) algorithm to automatically calculate the exact planning horizon needed. We then "flattened" the timeline: a technician working 5 days is dynamically split into 5 "Virtual Technicians" (tech-days), and a location needing two visits becomes 2 "Virtual Locations." This reduces a highly complex, multi-dimensional scheduling problem into a streamlined 2D model, allowing the solver to process it with maximum performance.

3. **Resilience to Imperfect Data: Dynamic Excel Parsing**  
   We completely abandoned rigid, hardcoded column indexes. Instead, we developed a smart, dynamic header mapper. It scans multi-level headers, normalizes the text (ignoring case, spaces, and special characters), and automatically maps the data to the correct system variables. Even if a user types "site addresss" with a typo or moves a column entirely, the system remains stable and processes the file correctly.

4. **Geodata Management: Caching & Cost Minimization**  
   We implemented a robust geo-caching layer. Before making a paid request to the AWS API, the system checks our internal database for existing coordinates. This drastically reduces AWS Lambda execution time (lowering compute costs) and drops third-party API expenses to near-zero for recurring client locations.

5. **Graceful Degradation (Soft vs. Hard Constraints)**  
   Instead of the algorithm failing and throwing a generic "Solution not found" error, we engineered the constraints using "active variables." If a location cannot be visited due to strict time or skill constraints, the algorithm safely "drops" it (taking a mathematical penalty) and continues building the optimal route for the remaining sites. As a result, the business always receives a highly optimized, usable schedule alongside a clear, actionable list of unassigned tasks, ensuring zero operational downtime.
