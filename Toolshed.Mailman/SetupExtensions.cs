using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using Microsoft.Extensions.Options;

namespace Toolshed.Mailman;

/// <summary>
/// Extension methods for configuring Mailman services in the dependency injection container
/// </summary>
public static class SetupExtensions
{
    /// <summary>
    /// Adds Mailman services to the dependency injection container using a configuration action
    /// </summary>
    /// <param name="services">The service collection to add services to</param>
    /// <param name="configureOptions">An action to configure the MailmanSettings</param>
    public static IServiceCollection AddMailman(this IServiceCollection services, Action<MailmanSettings> configureOptions)
    {
        var m = new MailmanSettings();
        configureOptions.Invoke(m);
        return services.AddMailman(m);
    }

    /// <summary>
    /// Adds Mailman services to the dependency injection container using configuration from IConfiguration.
    /// Expects a "MailmanSettings" section in the configuration.
    /// </summary>
    /// <param name="services">The service collection to add services to</param>
    /// <param name="config">The configuration containing a "MailmanSettings" section</param>
    public static IServiceCollection AddMailman(this IServiceCollection services, IConfiguration config)
    {
        return services.AddMailman(config.GetSection("MailmanSettings"));
    }

    /// <summary>
    /// Adds Mailman services to the dependency injection container using a specific configuration section
    /// </summary>
    /// <param name="services">The service collection to add services to</param>
    /// <param name="configSection">The configuration section containing Mailman settings</param>
    public static IServiceCollection AddMailman(this IServiceCollection services, IConfigurationSection configSection)
    {
        var m = new MailmanSettings();
        configSection.Bind(m);
        return services.AddMailman(m);
    }

    /// <summary>
    /// Adds Mailman services to the dependency injection container using IConfiguration with additional configuration action.
    /// Settings from configuration are applied first, then the action is invoked for further customization.
    /// </summary>
    /// <param name="services">The service collection to add services to</param>
    /// <param name="config">The configuration containing a "MailmanSettings" section</param>
    /// <param name="configureOptions">An action to further configure or override the MailmanSettings</param>
    public static IServiceCollection AddMailman(this IServiceCollection services, IConfiguration config, Action<MailmanSettings> configureOptions)
    {
        return services.AddMailman(config.GetSection("MailmanSettings"), configureOptions);
    }

    /// <summary>
    /// Adds Mailman services to the dependency injection container using a configuration section with additional configuration action.
    /// Settings from the section are applied first, then the action is invoked for further customization.
    /// </summary>
    /// <param name="services">The service collection to add services to</param>
    /// <param name="configSection">The configuration section containing Mailman settings</param>
    /// <param name="configureOptions">An action to further configure or override the MailmanSettings</param>
    public static IServiceCollection AddMailman(this IServiceCollection services, IConfigurationSection configSection, Action<MailmanSettings> configureOptions)
    {
        var m = new MailmanSettings();
        configSection.Bind(m);
        configureOptions.Invoke(m);
        return services.AddMailman(m);
    }

    /// <summary>
    /// Adds Mailman services to the dependency injection container using a pre-configured MailmanSettings instance.
    /// Registers MailmanSettings as a singleton and MailmanService as transient.
    /// </summary>
    /// <param name="services">The service collection to add services to</param>
    /// <param name="settings">The configured MailmanSettings instance</param>
    /// <param name="useTransient">Whether to register MailmanService as transient (true) or scoped (false)</param>
    public static IServiceCollection AddMailman(this IServiceCollection services, MailmanSettings settings, bool useTransient = false)
    {
        services.AddSingleton(settings);

        if(useTransient)
        {
            services.AddTransient<MailmanService>();
        }
        else
        {
            services.AddScoped<MailmanService>();
        }
        
        return services;
    }
}
