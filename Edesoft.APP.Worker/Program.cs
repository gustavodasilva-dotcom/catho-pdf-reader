using Edesoft.APP.Abstractions.Excel;
using Edesoft.APP.Abstractions.IO;
using Edesoft.APP.Abstractions.Pdf;
using Edesoft.APP.Application.Excel;
using Edesoft.APP.Application.IO;
using Edesoft.APP.Application.Pdf;
using Edesoft.APP.Infrastructure.Configuration;
using Edesoft.APP.Worker;

IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<Worker>();
        services.AddSingleton<IIOApplication, IOApplication>();
        services.AddSingleton<IPdfApplication, PdfApplication>();
        services.AddSingleton<IExcelApplication, ExcelApplication>();
    })
    .Build();

var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();

Settings.Configuration = configuration;

host.Run();
