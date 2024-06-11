using ExcelCellCounter.Components;
using Microsoft.AspNetCore.ResponseCompression;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddResponseCompression(options =>
{
	options.Providers.Add<BrotliCompressionProvider>();
    options.MimeTypes = ResponseCompressionDefaults.MimeTypes.Concat(new[] { "application/octet-stream" });
    options.EnableForHttps = true;
});


var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.Use(async (context, next) =>
{
	context.Response.Headers.Append("Content-Security-Policy", "default-src 'self'; script-src 'self' 'nonce-somestring' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; base-uri 'self';");
	await next();
});


app.UseResponseCompression();
app.UseStaticFiles(new StaticFileOptions
{
	OnPrepareResponse = ctx =>
	{
		// Cache static files for 30 days
		ctx.Context.Response.Headers.Append("Cache-Control", "public,max-age=31536000");
	}
});

app.UseHttpsRedirection();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
