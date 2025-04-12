using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using Microsoft.Extensions.Options;

namespace Toolshed.Mailman
{
    public static class SetupExtensions
    {
        public static void AddMailman(this IServiceCollection services, Action<MailmanSettings> configureOptions)
        {
            var m = new MailmanSettings();
            configureOptions.Invoke(m);
            services.AddMailman(m);
        }
        public static void AddMailman(this IServiceCollection services, IConfiguration config)
        {
            var m = new MailmanSettings();
            config.GetSection("MailmanSettings").Bind(m);
            services.AddMailman(m);
        }
        public static void AddMailman(this IServiceCollection services, IConfigurationSection configSection)
        {
            var m = new MailmanSettings();
            configSection.Bind(m);
            services.AddMailman(m);
        }
        public static void AddMailman(this IServiceCollection services, IConfigurationRoot config)
        {
            var m = new MailmanSettings();
            config.GetSection("MailmanSettings").Bind(m);
            services.AddMailman(m);
        }


        public static void AddMailman(this IServiceCollection services, IConfiguration config, Action<MailmanSettings> configureOptions)
        {
            var m = new MailmanSettings();
            config.GetSection("MailmanSettings").Bind(m);
            configureOptions.Invoke(m);
            services.AddMailman(m);
        }
        public static void AddMailman(this IServiceCollection services, IConfigurationSection configSection, Action<MailmanSettings> configureOptions)
        {
            var m = new MailmanSettings();
            configSection.Bind(m);
            configureOptions.Invoke(m);
            services.AddMailman(m);
        }
        public static void AddMailman(this IServiceCollection services, IConfigurationRoot config, Action<MailmanSettings> configureOptions)
        {
            var m = new MailmanSettings();
            config.GetSection("MailmanSettings").Bind(m);
            configureOptions.Invoke(m);
            services.AddMailman(m);
        }

        public static void AddMailman(this IServiceCollection services, MailmanSettings settings)
        {
            services.AddSingleton(settings);
            services.AddTransient<MailmanService>();
        }
    }
}
