using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Toolshed.Mailman
{
    public static class SetupExtensions
    {
        public static void UseMailman(this IServiceCollection services, Action<MailmanSettings> configureOptions)
        {
            var m = new MailmanSettings();
            configureOptions.Invoke(m);
            services.UseMailman(m);
        }
        public static void UseMailman(this IServiceCollection services, IConfiguration config)
        {
            var m = new MailmanSettings();
            config.GetSection("MailmanSettings").Bind(m);
            services.UseMailman(m);
        }
        public static void UseMailman(this IServiceCollection services, IConfigurationSection configSection)
        {
            var m = new MailmanSettings();
            configSection.Bind(m);
            services.UseMailman(m);
        }
        public static void UseMailman(this IServiceCollection services, MailmanSettings settings)
        {
            services.AddSingleton(settings);
            services.AddTransient<MailmanService>();
            services.AddScoped<ViewRenderService>();
        }

    }
}
