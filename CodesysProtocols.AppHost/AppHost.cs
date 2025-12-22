var builder = DistributedApplication.CreateBuilder(args);

builder.AddProject<Projects.CodesysProtocols_Blazor>("codesysprotocols-blazor");

builder.Build().Run();
