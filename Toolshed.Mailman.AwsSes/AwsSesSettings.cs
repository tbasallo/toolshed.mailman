using Amazon;

namespace Toolshed.Mailman.AwsSes;

/// <summary>
/// Configuration settings for the Amazon SES (Simple Email Service) email service.
/// </summary>
public class AwsSesSettings
{
    /// <summary>
    /// The AWS access key id used to authenticate against the SES API. When left null, the AWS SDK
    /// default credential resolution is used (environment variables, shared credentials file,
    /// IAM role, etc.), which is the recommended approach when running on AWS infrastructure.
    /// </summary>
    public string? AccessKeyId { get; set; }

    /// <summary>
    /// The AWS secret access key that pairs with <see cref="AccessKeyId"/>. When left null, the AWS SDK
    /// default credential resolution is used.
    /// </summary>
    public string? SecretAccessKey { get; set; }

    /// <summary>
    /// An optional AWS session token to use alongside temporary credentials.
    /// </summary>
    public string? SessionToken { get; set; }

    /// <summary>
    /// The AWS region system name that hosts the SES endpoint (for example "us-east-1"). Required when
    /// credentials are supplied explicitly and no <see cref="ServiceUrl"/> override is provided.
    /// </summary>
    public string? Region { get; set; }

    /// <summary>
    /// An optional service URL override for the SES endpoint. Useful for testing against a local mock or
    /// a custom/VPC endpoint. When set, it takes precedence over <see cref="Region"/>.
    /// </summary>
    public string? ServiceUrl { get; set; }

    /// <summary>
    /// The default sender email address used when From is not explicitly set. The address must be a
    /// verified identity in the SES account/region.
    /// </summary>
    public string? FromAddress { get; set; }

    /// <summary>
    /// The default sender display name used when From is not explicitly set.
    /// </summary>
    public string? FromDisplayName { get; set; }

    /// <summary>
    /// The name of the SES configuration set to apply to sent messages. Configuration sets enable event
    /// publishing (bounces, complaints, deliveries) and dedicated IP pool selection.
    /// </summary>
    public string? ConfigurationSetName { get; set; }

    /// <summary>
    /// The address SES should use for the return-path (envelope MAIL FROM) of a message. Optional.
    /// </summary>
    public string? FeedbackForwardingEmailAddress { get; set; }

    /// <summary>
    /// A comma delimited list of message tags (SES "tags") applied to every message. Each entry is a
    /// name that is sent to SES with a value of "true". This mirrors the categories concept in the SMTP
    /// based library and is useful for filtering CloudWatch/event metrics.
    /// </summary>
    public string? Categories { get; set; }

    /// <summary>
    /// The timeout in milliseconds for SES API calls. If null, the AWS SDK default timeout is used.
    /// </summary>
    public int? Timeout { get; set; }

    /// <summary>
    /// Resolves the configured <see cref="RegionEndpoint"/> from <see cref="Region"/>, or null when no
    /// region is configured.
    /// </summary>
    public RegionEndpoint? GetRegionEndpoint()
    {
        return string.IsNullOrWhiteSpace(Region) ? null : RegionEndpoint.GetBySystemName(Region);
    }
}
