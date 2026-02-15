using Microsoft.Extensions.Logging;
using CommunityToolkit.Maui;
using CommunityToolkit.Maui.Storage;


namespace ShaiMauiBlazorWindows;

public static class MauiProgram
{
	public static MauiApp CreateMauiApp()
	{
		var builder = MauiApp.CreateBuilder();
		builder
			.UseMauiApp<App>()
            .UseMauiCommunityToolkit()
			.ConfigureFonts(fonts =>
			{
				fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
			});

		builder.Services.AddMauiBlazorWebView();

        // Register Services
        builder.Services.AddSingleton<ShaiMauiExcelToSql.Services.DatabaseService>();
        builder.Services.AddSingleton<ShaiMauiExcelToSql.Services.QueryHistoryService>();
        builder.Services.AddSingleton<ShaiMauiExcelToSql.Services.ExcelExportSettingsService>();
        builder.Services.AddSingleton<CommunityToolkit.Maui.Storage.IFolderPicker>(CommunityToolkit.Maui.Storage.FolderPicker.Default);

#if DEBUG
		builder.Services.AddBlazorWebViewDeveloperTools();
		builder.Logging.AddDebug();
#endif
		return builder.Build();
	}
}
