using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using MailKit.Net.Smtp;

using MimeKit;
using MimeKit.IO;

namespace Toolshed.Mailman;

/// <summary>
/// Base service class for Mailman that provides common functionality
/// </summary>
public class MailmanBaseService
{
    /// <summary>
    /// Saves email to a local directory
    /// </summary>
    /// <param name="message"></param>
    /// <param name="pickupDirectory"></param>
    public static void SaveToPickupDirectory(MimeMessage message, string pickupDirectory)
    {
        do
        {
            // Generate a random file name to save the message to.
            var path = Path.Combine(pickupDirectory, Guid.NewGuid().ToString() + ".eml");
            Stream stream;

            try
            {
                // Attempt to create the new file.
                stream = File.Open(path, FileMode.CreateNew);
            }
            catch (IOException)
            {
                // If the file already exists, try again with a new Guid.
                if (File.Exists(path))
                    continue;

                // Otherwise, fail immediately since it probably means that there is
                // no graceful way to recover from this error.
                throw;
            }

            try
            {
                using (stream)
                {
                    // IIS pickup directories expect the message to be "byte-stuffed"
                    // which means that lines beginning with "." need to be escaped
                    // by adding an extra "." to the beginning of the line.
                    //
                    // Use an SmtpDataFilter "byte-stuff" the message as it is written
                    // to the file stream. This is the same process that an SmtpClient
                    // would use when sending the message in a `DATA` command.
                    using var filtered = new FilteredStream(stream);
                    filtered.Add(new SmtpDataFilter());

                    // Make sure to write the message in DOS (<CR><LF>) format.
                    var options = FormatOptions.Default.Clone();
                    options.NewLineFormat = NewLineFormat.Dos;

                    message.WriteTo(options, filtered);
                    filtered.Flush();
                    return;
                }
            }
            catch
            {
                // An exception here probably means that the disk is full.
                //
                // Delete the file that was created above so that incomplete files are not
                // left behind for IIS to send accidentally.
                File.Delete(path);
                throw;
            }
        } while (true);
    }

    /// <summary>
    /// Adds the X-SMTPAPI header with a properly formatted JSON array of categories when any are configured.
    /// The instance <see cref="Categories"/> takes precedence over categories provided via settings.
    /// </summary>
    public void ApplyCategories(MimeMessage mailMessage, string? categories)
    {
        if (string.IsNullOrWhiteSpace(categories))
        {
            return;
        }

        // Remove any previously applied header so re-sending the same message does not duplicate it.
        mailMessage.Headers.RemoveAll("X-SMTPAPI");

        if (string.IsNullOrWhiteSpace(categories))
        {
            return;
        }

        var items = categories
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (items.Length == 0)
        {
            return;
        }

        var payload = System.Text.Json.JsonSerializer.Serialize(new { category = items });
        mailMessage.Headers.Add("X-SMTPAPI", payload);
    }
}
