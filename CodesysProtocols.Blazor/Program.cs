using CodesysProtocols.Blazor.Components;
using CodesysProtocols.Blazor.Services;
using CodesysProtocols.Model;
using Microsoft.FluentUI.AspNetCore.Components;

var builder = WebApplication.CreateBuilder(args);

builder.AddServiceDefaults();

// Add services to the container.
builder.Services.AddScoped<IIec608705ConverterService, Iec608705ConverterService>();
builder.Services.AddScoped<IIec608705ExcelService, Iec608705ExcelService>();
builder.Services.AddScoped<IIec608705ExcelWorkbookValidationService, Iec608705ExcelWorkbookValidationService>();
builder.Services.AddScoped<Iec608705ExcelWorkbookValidation>();
builder.Services.AddScoped<Iec608705Converter>();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddFluentUIComponents();

var app = builder.Build();

app.MapDefaultEndpoints();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
