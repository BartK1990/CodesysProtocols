using CodesysProtocols.Blazor.Services.Codesys23;
using CodesysProtocols.DataAccess.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.FluentUI.AspNetCore.Components;
using Microsoft.JSInterop;
using System.Xml.Linq;

namespace CodesysProtocols.Blazor.Components.Pages.Codesys23;

public partial class Iec608705Converter
{
    [Inject] private IJSRuntime JS { get; set; } = default!;

    [Inject] private IIec608705ConverterService IecConverterService { get; set; } = default!;

    [Inject] private IIec608705ExcelService IecExcelService { get; set; } = default!;

    [Inject] private IIec608705ExcelWorkbookValidationService IecExcelWorkbookValidationService { get; set; } = default!;

    [Inject] private IIec608705CounterService Iec608705CounterService { get; set; } = default!;

    private int ProgressPercent = 0;
    private FluentInputFileEventArgs[] Files = [];
    private string LoggerText { get; set; } = string.Empty;
    private bool ProtocolConfigurationLoading { get; set; }
    private bool XmlConfigurationSaving { get; set; }
    private string? ProtocolConfigurationName { get; set; }
    private long ToExcelCounter { get; set; }
    private long ToXmlCounter { get; set; }

    // Generated outputs after conversion
    private byte[]? _generatedExcelBytes;
    private byte[]? _generatedXmlBytes;

    private bool CanDownloadExcel => _generatedExcelBytes is not null;
    private bool CanDownloadXml => _generatedXmlBytes is not null;

    protected override async Task OnInitializedAsync() 
    {
        ToExcelCounter = await Iec608705CounterService.GetCounterValueAsync(DataAccess.Enums.Codesys23Iec608705CounterType.ToExcel);
        ToXmlCounter = await Iec608705CounterService.GetCounterValueAsync(DataAccess.Enums.Codesys23Iec608705CounterType.ToXml);
    }

    private async Task OnCompletedFileAsync(IEnumerable<FluentInputFileEventArgs> files)
    {
        Files = files.ToArray();
        foreach (var file in Files)
        {
            await HandleFileAsync(file);
        }

        ProgressPercent = 0;
    }

    private async Task HandleFileAsync(FluentInputFileEventArgs file)
    {
        ProtocolConfigurationLoading = true;
        ResetOutputs();
        StateHasChanged();

        if(file.ErrorMessage is not null)
        {
            LogError(file.ErrorMessage);
            return;
        }

        if (file.Stream is null)
        {
            LogError("No file stream available");
            return;
        }
        using var fileStream = new MemoryStream();
        await file.Stream.CopyToAsync(fileStream);
        try
        {
            ProtocolConfigurationName = file.Name;
            var ext = Path.GetExtension(file.Name);
            if (file.Stream is null)
            {
                LogError("No file stream available");
                return;
            }
            switch (ext)
            {
                case ".xlsx":
                case ".xlsm":
                case ".xls":

                    string[] sheetsNames = await IecExcelService.GetExcelSheetNamesAsync(fileStream);
                    if (!await IecExcelWorkbookValidationService.CheckIfContainsValidSheetsAsync(sheetsNames))
                    {
                        LogError("Excel does not contain all required sheet names");
                        return;
                    }

                    var tablesFromExcel = await IecExcelService.ReadTablesFromExcelAsync(fileStream);
                    if (tablesFromExcel is null)
                    {
                        LogError("Excel does not contain all required sheet names");
                        break;
                    }

                    var xmlDoc = await IecConverterService.DataToXmlAsync(tablesFromExcel);
                    using (var ms = new MemoryStream())
                    {
                        xmlDoc.Save(ms);
                        _generatedXmlBytes = ms.ToArray();
                    }

                    ToXmlCounter++;
                    Log("Protocol configuration Excel loaded");
                    break;

                case ".xml":
                    fileStream.Position = 0;
                    XDocument xml = XDocument.Load(fileStream);
                    var tablesFromXml = await IecConverterService.XmlToDataAsync(xml);
                    using (var ms = new MemoryStream())
                    {
                        await IecExcelService.GetExcelFromTablesAsync(ms, tablesFromXml);
                        _generatedExcelBytes = ms.ToArray();
                    }

                    ToExcelCounter++;
                    Log("Protocol configuration XML file loaded");
                    break;

                default:
                    Log("Drop something else! Wrong file type");
                    ProtocolConfigurationName = null;
                    break;
            }
        }
        catch (Exception e)
        {
            LogError($"File loading failure. Error: {e.Message}");
            ProtocolConfigurationName = null;
        }
        finally
        {
            ProtocolConfigurationLoading = false;
            StateHasChanged();
        }
    }

    private void ResetOutputs()
    {
        _generatedExcelBytes = null;
        _generatedXmlBytes = null;
    }

    private void Log(string log)
    {
        var time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var logStr = $"{time} | {log}";
        LoggerText += $"{logStr}{Environment.NewLine}";
    }

    private void LogError(string log) => Log($"ERROR: {log}");

    private async Task DownloadExcel()
    {
        if (_generatedExcelBytes is null)
        {
            Log("Load something first");
            return;
        }

        try
        {
            XmlConfigurationSaving = true;
            await TriggerBrowserDownloadAsync(_generatedExcelBytes, MakeOutputFileName(".xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            Log("Protocol configuration Excel file prepared for download");
        }
        catch (Exception e)
        {
            LogError($"File saving failure. Error: {e.Message}");
        }
        finally
        {
            XmlConfigurationSaving = false;
        }
    }

    private async Task DownloadXml()
    {
        if (_generatedXmlBytes is null)
        {
            Log("Load something first");
            return;
        }

        try
        {
            XmlConfigurationSaving = true;
            await TriggerBrowserDownloadAsync(_generatedXmlBytes, MakeOutputFileName(".xml"), "application/xml");
            Log("Protocol configuration XML file prepared for download");
        }
        catch (Exception e)
        {
            LogError($"File saving failure. Error: {e.Message}");
        }
        finally
        {
            XmlConfigurationSaving = false;
        }
    }

    private string MakeOutputFileName(string ext)
    {
        var baseName = string.IsNullOrWhiteSpace(ProtocolConfigurationName) ? "protocol" : Path.GetFileNameWithoutExtension(ProtocolConfigurationName);
        return $"{baseName}{ext}";
    }

    private async Task TriggerBrowserDownloadAsync(byte[] data, string name, string contentType)
    {
        var base64 = Convert.ToBase64String(data);
        await JS.InvokeVoidAsync("downloadFile", name, contentType, base64);
    }

    private void ClearLog()
    {
        LoggerText = string.Empty;
    }
}