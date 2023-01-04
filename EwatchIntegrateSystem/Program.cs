using EwatchIntegrateSystem.Hubs;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    var xmlFilename = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
    options.IncludeXmlComments(Path.Combine(AppContext.BaseDirectory, xmlFilename));
});
builder.Services.AddSignalR();//加入SignalR功能
builder.Services.AddCors();
//builder.WebHost.ConfigureKestrel(serveroption =>
//{
//    serveroption.ListenAnyIP(200);
//});
var app = builder.Build();

// Configure the HTTP request pipeline.
app.UseSwagger();
app.UseSwaggerUI();
if (app.Environment.IsDevelopment())
{
}
app.UseCors(builder =>
{
    builder.AllowAnyHeader().AllowAnyMethod().SetIsOriginAllowed(_ => true).AllowCredentials();
});
app.UseAuthorization();

app.MapControllers();

app.MapHub<SystemHub>("/systemHub");//開啟SignalR_Hub
app.MapHub<InformationDataHub>("/informationHub");//開啟SignalR_Hub

app.Run();
