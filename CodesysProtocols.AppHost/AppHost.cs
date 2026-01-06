var builder = DistributedApplication.CreateBuilder(args);

builder.AddDockerComposeEnvironment("compose")
    .WithProperties(env =>
    {
        env.DefaultNetworkName = "codesysprotocols_net";
    })
    .WithDashboard(dashboard =>
    {
        dashboard.WithHostPort(8080)
                .WithForwardedHeaders(enabled: true);
    });

var database = builder.AddSqlServer("database")
    .AddDatabase("CodesysProtocols");

builder.AddProject<Projects.CodesysProtocols_Blazor>("codesysprotocols-blazor")
    .WithReference(database)
    .WaitFor(database)
    .PublishAsDockerComposeService((serviceResource, service) => 
    {
        service.ContainerName = serviceResource.Name;
    });

builder.Build().Run();
