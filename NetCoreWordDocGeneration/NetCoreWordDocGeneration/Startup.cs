using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Net.Http.Headers;
using NetCoreWordDocGeneration.Util;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace NetCoreWordDocGeneration
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            //CORS Add service and create Policy with options
            services.AddCors(options =>
            {
                options.AddPolicy("CorsPolicy",
                    builder => builder.AllowAnyOrigin()
                    .AllowAnyMethod()
                    .AllowAnyHeader()
                    .AllowCredentials());
            });

            services.AddMvc(config =>
            {
                config.OutputFormatters.Add(new WordOutputFormatter());
                config.FormatterMappings.SetMediaTypeMappingForFormat(
                  "docx", MediaTypeHeaderValue.Parse("application/ms-word"));

            }).AddJsonOptions(options =>
            {
                options.SerializerSettings.ContractResolver = new CamelCasePropertyNamesContractResolver();
                options.SerializerSettings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;
                options.SerializerSettings.DefaultValueHandling = DefaultValueHandling.Include;
                options.SerializerSettings.NullValueHandling = NullValueHandling.Ignore;
            });

            //add swagger (API documentation)
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new Swashbuckle.AspNetCore.Swagger.Info { Title = "My App!", Version = "v1" });

            });
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseCors("CorsPolicy");

            app.UseMvc();

            app.UseSwagger(c => { c.RouteTemplate = "api/swagger/{documentName}/swagger.json"; });
            app.UseSwaggerUI(c => { c.SwaggerEndpoint("v1/swagger.json", "My App Api"); c.RoutePrefix = "api/swagger"; });
        }
    }
}
