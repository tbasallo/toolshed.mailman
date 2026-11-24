using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Toolshed.Mailman.AwsSes;

/// <summary>
/// Extension methods for configuring the Amazon SES Mailman services in the dependency injection container.
/// </summary>
public static class SetupExtensions
{
    /// <summary>
    /// Adds the SES Mailman services to the dependency injection container using a configuration action.
    /// </summary>
    /// <param name="services">The service collection to add services to.</param>
    /// <param name="configureOptions">An action to configure the <see cref="AwsSesSettings"/>.</param>
    public static IServiceCollection AddAwsSesMailman(this IServiceCollection services, Action<AwsSesSettings> configureOptions)
    {
        var settings = new AwsSesSettings();
        configureOptions.Invoke(settings);
        return services.AddAwsSesMailman(settings);
    }

    /// <summary>
    /// Adds the SES Mailman services to the dependency injection container using configuration from IConfiguration.
    /// Expects an "AwsSesSettings" section in the configuration.
    /// </summary>
    /// <param name="services">The service collection to add services to.</param>
    /// <param name="config">The configuration containing an "AwsSesSettings" section.</param>
    public static IServiceCollection AddAwsSesMailman(this IServiceCollection services, IConfiguration config)
    {
        return services.AddAwsSesMailman(config.GetSection("AwsSesSettings"));
    }

    /// <summary>
    /// Adds the SES Mailman services to the dependency injection container using a specific configuration section.
    /// </summary>
    /// <param name="services">The service collection to add services to.</param>
    /// <param name="configSection">The configuration section containing SES settings.</param>
    public static IServiceCollection AddAwsSesMailman(this IServiceCollection services, IConfigurationSection configSection)
    {
        var settings = new AwsSesSettings();
        configSection.Bind(settings);
        return services.AddAwsSesMailman(settings);
    }

    /// <summary>
    /// Adds the SES Mailman services to the dependency injection container using IConfiguration with an
    /// additional configuration action. Settings from configuration are applied first, then the action is invoked.
    /// </summary>
    /// <param name="services">The service collection to add services to.</param>
    /// <param name="config">The configuration containing an "AwsSesSettings" section.</param>
    /// <param name="configureOptions">An action to further configure or override the <see cref="AwsSesSettings"/>.</param>
    public static IServiceCollection AddAwsSesMailman(this IServiceCollection services, IConfiguration config, Action<AwsSesSettings> configureOptions)
    {
        return services.AddAwsSesMailman(config.GetSection("AwsSesSettings"), configureOptions);
    }

    /// <summary>
    /// Adds the SES Mailman services to the dependency injection container using a configuration section with
    /// an additional configuration action. Settings from the section are applied first, then the action is invoked.
    /// </summary>
    /// <param name="services">The service collection to add services to.</param>
    /// <param name="configSection">The configuration section containing SES settings.</param>
    /// <param name="configureOptions">An action to further configure or override the <see cref="AwsSesSettings"/>.</param>
    public static IServiceCollection AddAwsSesMailman(this IServiceCollection services, IConfigurationSection configSection, Action<AwsSesSettings> configureOptions)
    {
        var settings = new AwsSesSettings();
        configSection.Bind(settings);
        configureOptions.Invoke(settings);
        return services.AddAwsSesMailman(settings);
    }

    /// <summary>
    /// Adds the SES Mailman services to the dependency injection container using a pre-configured
    /// <see cref="AwsSesSettings"/> instance. Registers the settings as a singleton and
    /// <see cref="AwsSesMailmanService"/> as scoped.
    /// </summary>
    /// <param name="services">The service collection to add services to.</param>
    /// <param name="settings">The configured <see cref="AwsSesSettings"/> instance.</param>
    public static IServiceCollection AddAwsSesMailman(this IServiceCollection services, AwsSesSettings settings)
    {
        services.AddSingleton(settings);
        services.AddScoped<AwsSesMailmanService>();

        return services;
    }
}
