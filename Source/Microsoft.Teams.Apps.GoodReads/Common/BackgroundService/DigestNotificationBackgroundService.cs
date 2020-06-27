// <copyright file="DigestNotificationBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.BackgroundService
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Rest.Azure;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.WindowsAzure.Storage;

    /// <summary>
    /// This class inherits IHostedService and implements the methods related to background tasks for sending Weekly/Monthly notifications.
    /// </summary>
    public class DigestNotificationBackgroundService : BackgroundService
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<DigestNotificationBackgroundService> logger;

        /// <summary>
        /// Instance of notification helper which helps in sending notifications.
        /// </summary>
        private readonly IDigestNotificationHelper digestNotificationHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="DigestNotificationBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to sending notification tasks.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="notificationHelper">Helper to send notification in channels.</param>
        public DigestNotificationBackgroundService(
            ILogger<DigestNotificationBackgroundService> logger,
            IDigestNotificationHelper notificationHelper)
        {
            this.logger = logger;
            this.digestNotificationHelper = notificationHelper;
        }

        /// <summary>
        ///  This method is called when the Microsoft.Extensions.Hosting.IHostedService starts.
        ///  The implementation should return a task that represents the lifetime of the long
        ///  running operation(s) being performed.
        /// </summary>
        /// <param name="stoppingToken">Triggered when Microsoft.Extensions.Hosting.IHostedService.StopAsync(System.Threading.CancellationToken) is called.</param>
        /// <returns>A System.Threading.Tasks.Task that represents the long running operations.</returns>
        protected async override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    var currentDateTime = DateTimeOffset.UtcNow.AddDays(1);
                    this.logger.LogInformation($"Notification Hosted Service is running at: {currentDateTime}.");

                    if (currentDateTime.DayOfWeek == DayOfWeek.Monday)
                    {
                        this.logger.LogInformation($"Monday of the month: {currentDateTime} and sending the notification.");
                        DateTime fromDate = currentDateTime.AddDays(-7).Date;
                        DateTime toDate = currentDateTime.Date;

                        this.logger.LogInformation("Notification task queued for sending weekly notification.");
                        await this.digestNotificationHelper.SendNotificationInChannelAsync(fromDate, toDate, Constants.WeeklyDigest); // Send the notifications
                    }

                    // Send digest notification if it's the 1st day of the Month.
                    if (currentDateTime.Day == 1)
                    {
                        this.logger.LogInformation($"First day of the month: {currentDateTime} and sending the notification.");
                        DateTime fromDate = currentDateTime.AddMonths(-1).Date;
                        DateTime toDate = currentDateTime.Date;

                        this.logger.LogInformation("Notification task queued for sending monthly notification.");
                        await this.digestNotificationHelper.SendNotificationInChannelAsync(fromDate, toDate, Constants.MonthlyDigest); // Send the notifications
                    }
                }
                catch (CloudException ex)
                {
                    this.logger.LogError(ex, $"Error occurred while accessing search service: {ex.Message} at: {DateTimeOffset.UtcNow}");
                }
                catch (StorageException ex)
                {
                    this.logger.LogError(ex, $"Error occurred while accessing storage: {ex.Message} at: {DateTimeOffset.UtcNow}");
                }
#pragma warning disable CA1031 // Catching general exceptions that might arise during execution to avoid blocking next run.
                catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions that might arise during execution to avoid blocking next run.
                {
                    this.logger.LogError(ex, "Error occurred while running digest notification service.");
                }

                await Task.Delay(TimeSpan.FromDays(1), stoppingToken);
            }
        }
    }
}
