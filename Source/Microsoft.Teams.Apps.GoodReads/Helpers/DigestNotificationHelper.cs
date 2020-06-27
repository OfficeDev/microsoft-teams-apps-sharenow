// <copyright file="DigestNotificationHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoodReads.Cards;
    using Microsoft.Teams.Apps.GoodReads.Common;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.Teams.Apps.GoodReads.Models.Configuration;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// A class that handles sending notification to different channels.
    /// </summary>
    public class DigestNotificationHelper : IDigestNotificationHelper
    {
        /// <summary>
        /// default value for channel activity to send notifications.
        /// </summary>
        private const string Channel = "msteams";

        /// <summary>
        /// Max post count for list card.
        /// </summary>
        private const int ListCardPostCount = 15;

        /// <summary>
        /// Retry policy with jitter.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private readonly AsyncRetryPolicy retryPolicy;

        /// <summary>
        /// Provider to fetch tab configuration for team.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// Helper for storing channel details to azure table storage for sending notification.
        /// </summary>
        private readonly ITeamPreferenceStorageProvider teamPreferenceStorageProvider;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<DigestNotificationHelper> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Represents a set of key/value application configuration properties for bot.
        /// </summary>
        private readonly IOptions<BotSettings> botOptions;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<GoodReadsActivityHandlerOptions> options;

        /// <summary>
        /// Instance of Search service for working with storage.
        /// </summary>
        private readonly IPostSearchService teamPostSearchService;

        /// <summary>
        /// Instance of team post storage helper to update post and get information of posts.
        /// </summary>
        private readonly IPostStorageHelper teamPostStorageHelper;

        /// <summary>
        /// Card post type images pair.
        /// </summary>
        private readonly Dictionary<int, string> cardPostTypePair = new Dictionary<int, string>();

        /// <summary>
        /// Initializes a new instance of the <see cref="DigestNotificationHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="botOptions">A set of key/value application configuration properties for bot.</param>
        /// <param name="adapter">Bot adapter.</param>
        /// <param name="teamPreferenceStorageProvider">Storage provider for team preference.</param>
        /// <param name="teamPostSearchService">The team post search service dependency injection.</param>
        /// <param name="teamPostStorageHelper">Team post storage helper dependency injection.</param>
        /// <param name="teamTagStorageProvider">Provider to fetch tab configuration for team.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        public DigestNotificationHelper(
            ILogger<DigestNotificationHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<BotSettings> botOptions,
            IBotFrameworkHttpAdapter adapter,
            ITeamPreferenceStorageProvider teamPreferenceStorageProvider,
            IPostSearchService teamPostSearchService,
            IPostStorageHelper teamPostStorageHelper,
            ITeamTagStorageProvider teamTagStorageProvider,
            IOptions<GoodReadsActivityHandlerOptions> options)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.botOptions = botOptions ?? throw new ArgumentNullException(nameof(botOptions));
            this.adapter = adapter;
            this.teamPreferenceStorageProvider = teamPreferenceStorageProvider;
            this.teamPostSearchService = teamPostSearchService;
            this.teamTagStorageProvider = teamTagStorageProvider;
            this.teamPostStorageHelper = teamPostStorageHelper;
            this.options = options;
            this.retryPolicy = Policy.Handle<Exception>()
                .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(this.botOptions.Value.MedianFirstRetryDelay), this.botOptions.Value.RetryCount));
        }

        /// <summary>
        /// Send notification in channels on weekly or monthly basis as per the configured preference in different channels.
        /// Fetch data based on the date range and send it accordingly.
        /// </summary>
        /// <param name="startDate">Start date from which data should fetch.</param>
        /// <param name="endDate">End date till when data should fetch.</param>
        /// <param name="digestFrequency">Digest frequency text for notification like Monthly/Weekly.</param>
        /// <returns>A task that sends notification in channel.</returns>
        public async Task SendNotificationInChannelAsync(DateTime startDate, DateTime endDate, string digestFrequency)
        {
            this.logger.LogInformation($"Send notification Timer trigger function executed at: {DateTime.UtcNow}");

            var teamPosts = await this.teamPostSearchService.GetPostsAsync(PostSearchScope.FilterPostsAsPerDateRange, searchQuery: null, userObjectId: null);
            var filteredTeamPosts = this.teamPostStorageHelper.GetTeamPostsForDateRange(teamPosts, startDate, endDate);

            if (filteredTeamPosts.Any())
            {
                var teamPreferences = await this.teamPreferenceStorageProvider.GetTeamPreferencesByDigestFrequencyAsync(digestFrequency);
                var notificationCardTitle = digestFrequency == Constants.WeeklyDigest
                    ? this.localizer.GetString("NotificationCardWeeklyTitleText")
                    : this.localizer.GetString("NotificationCardMonthlyTitleText");

                if (teamPreferences != null)
                {
                    foreach (var teamPreference in teamPreferences)
                    {
                        var tagsFilteredData = this.GetDataAsPerTags(teamPreference, filteredTeamPosts);

                        if (tagsFilteredData.Any())
                        {
                            var notificationCard = DigestNotificationListCard.GetNotificationListCard(
                                tagsFilteredData,
                                this.localizer,
                                notificationCardTitle,
                                this.cardPostTypePair,
                                this.botOptions.Value.ManifestId,
                                this.options.Value.DiscoverTabEntityId,
                                this.options.Value.AppBaseUri);

                            var teamTabConfiguration = await this.teamTagStorageProvider.GetTeamTagAsync(teamPreference.TeamId);
                            if (teamTabConfiguration != null)
                            {
                                await this.SendCardToTeamAsync(teamPreference, notificationCard, teamTabConfiguration.ServiceUrl);
                            }
                        }
                    }
                }
                else
                {
                    this.logger.LogInformation("Unable to fetch team digest preferences.");
                }
            }
            else
            {
                this.logger.LogInformation($"There is no digest data available to send at this time range from: {0} till {1}", startDate, endDate);
            }
        }

        /// <summary>
        /// Send the given attachment to the specified team.
        /// </summary>
        /// <param name="teamPreferenceEntity">Team preference model object.</param>
        /// <param name="cardToSend">The attachment card to send.</param>
        /// <param name="serviceUrl">Service URL for a particular team.</param>
        /// <returns>A task that sends notification card in channel.</returns>
        private async Task SendCardToTeamAsync(
            TeamPreferenceEntity teamPreferenceEntity,
            Attachment cardToSend,
            string serviceUrl)
        {
            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
            string teamsChannelId = teamPreferenceEntity.TeamId;

            var conversationReference = new ConversationReference()
            {
                ChannelId = Channel,
                Bot = new ChannelAccount() { Id = $"28:{this.botOptions.Value.MicrosoftAppId}" },
                ServiceUrl = serviceUrl,
                Conversation = new ConversationAccount() { ConversationType = ConversationTypes.Channel, IsGroup = true, Id = teamsChannelId, TenantId = this.botOptions.Value.TenantId },
            };

            this.logger.LogInformation($"sending notification to channelId- {teamsChannelId}");

            // Retry it in addition to the original call.
            await this.retryPolicy.ExecuteAsync(async () =>
            {
                try
                {
                    await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                    this.botOptions.Value.MicrosoftAppId,
                    conversationReference,
                    async (conversationTurnContext, conversationCancellationToken) =>
                    {
                        await conversationTurnContext.SendActivityAsync(MessageFactory.Attachment(cardToSend));
                    },
                    CancellationToken.None);
                }
                catch (Exception ex)
                {
                    this.logger.LogError(ex, "Error while performing retry logic to send digest notification to channel.");
                    throw;
                }
            });
        }

        /// <summary>
        /// Get team posts as per configured tags for preference.
        /// </summary>
        /// <param name="teamPreferenceEntity">Team preference model object.</param>
        /// <param name="teamPosts">List of team posts.</param>
        /// <returns>List of team posts as per preference tags.</returns>
        private IEnumerable<PostEntity> GetDataAsPerTags(
            TeamPreferenceEntity teamPreferenceEntity,
            IEnumerable<PostEntity> teamPosts)
        {
            var filteredPosts = new List<PostEntity>();
            var preferenceTagList = teamPreferenceEntity.Tags.Split(";").Where(tag => !string.IsNullOrWhiteSpace(tag));
            bool isTagMatched = false;
            teamPosts = teamPosts.OrderByDescending(c => c.UpdatedDate);

            // Loop through the list of filtered posts.
            foreach (var teamPost in teamPosts)
            {
                // Split the comma separated post tags.
                var postTags = teamPost.Tags.Split(";").Where(tag => !string.IsNullOrWhiteSpace(tag));
                isTagMatched = false;

                // Loop through the list of preference tags.
                foreach (var preferenceTag in preferenceTagList)
                {
                    // Loop through the post tags.
                    foreach (var postTag in postTags)
                    {
                        // Check if the post tag and preference tag is same.
                        if (postTag.Trim() == preferenceTag.Trim())
                        {
                            // Set the flag to check the preference tag is present in post tag.
                            isTagMatched = true;
                            break; // break the loop to check for next preference tag with post tag.
                        }
                    }

                    if (isTagMatched && filteredPosts.Count < ListCardPostCount)
                    {
                        // If preference tag is present in post tag then add it in the list.
                        filteredPosts.Add(teamPost);
                        break; // break the inner loop to check for next post.
                    }
                }

                // Break the entire loop after getting top {ListCardPostCount} posts.
                if (filteredPosts.Count >= ListCardPostCount)
                {
                    break;
                }
            }

            return filteredPosts.Take(ListCardPostCount);
        }
    }
}
