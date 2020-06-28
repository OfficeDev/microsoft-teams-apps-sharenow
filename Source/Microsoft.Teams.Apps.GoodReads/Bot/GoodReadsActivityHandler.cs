// <copyright file="GoodReadsActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoodReads.Cards;
    using Microsoft.Teams.Apps.GoodReads.Common;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.Teams.Apps.GoodReads.Models.Configuration;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// This class is responsible for reacting to incoming events from Microsoft Teams sent from BotFramework.
    /// </summary>
    public sealed class GoodReadsActivityHandler : TeamsActivityHandler
    {
        /// <summary>
        /// Sets the height of the task module.
        /// </summary>
        private const int TaskModuleHeight = 460;

        /// <summary>
        /// Sets the width of the task module.
        /// </summary>
        private const int TaskModuleWidth = 600;

        /// <summary>
        /// Represents the close command for task module.
        /// </summary>
        private const string CloseCommand = "close";

        /// <summary>
        /// Submit preference command.
        /// </summary>
        private const string SubmitCommand = "submit";

        /// <summary>
        /// State management object for maintaining user conversation state.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<BotSettings> botOptions;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<GoodReadsActivityHandler> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Instance of Application Insights Telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Messaging Extension search helper for working with team posts data.
        /// </summary>
        private readonly IMessagingExtensionHelper messagingExtensionHelper;

        /// <summary>
        /// Instance of team preference storage helper.
        /// </summary>
        private readonly ITeamPreferenceStorageHelper teamPreferenceStorageHelper;

        /// <summary>
        /// Instance of team preference storage provider for team preferences.
        /// </summary>
        private readonly ITeamPreferenceStorageProvider teamPreferenceStorageProvider;

        /// <summary>
        /// Instance of team tags storage provider to configure team tags.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<GoodReadsActivityHandlerOptions> options;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoodReadsActivityHandler"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="options">>A set of key/value application configuration properties for activity handler.</param>
        /// <param name="messagingExtensionHelper">Messaging Extension helper dependency injection.</param>
        /// <param name="userState">State management object for maintaining user conversation state.</param>
        /// <param name="teamPreferenceStorageHelper">Team preference storage helper dependency injection.</param>
        /// <param name="teamPreferenceStorageProvider">Team preference storage provider dependency injection.</param>
        /// <param name="teamTagStorageProvider">Team tags storage provider dependency injection.</param>
        /// <param name="botOptions">A set of key/value application configuration properties for activity handler.</param>
        public GoodReadsActivityHandler(
            ILogger<GoodReadsActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<GoodReadsActivityHandlerOptions> options,
            IMessagingExtensionHelper messagingExtensionHelper,
            UserState userState,
            ITeamPreferenceStorageHelper teamPreferenceStorageHelper,
            ITeamPreferenceStorageProvider teamPreferenceStorageProvider,
            ITeamTagStorageProvider teamTagStorageProvider,
            IOptions<BotSettings> botOptions)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.messagingExtensionHelper = messagingExtensionHelper;
            this.userState = userState;
            this.teamPreferenceStorageHelper = teamPreferenceStorageHelper;
            this.teamPreferenceStorageProvider = teamPreferenceStorageProvider;
            this.teamTagStorageProvider = teamTagStorageProvider;
            this.botOptions = botOptions ?? throw new ArgumentNullException(nameof(botOptions));
        }

        /// <summary>
        /// Handles an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onturnasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        public override Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            this.RecordEvent(nameof(this.OnTurnAsync), turnContext);

            return base.OnTurnAsync(turnContext, cancellationToken);
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are removed from the conversation.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnConversationUpdateActivityAsync), turnContext);

                var activity = turnContext.Activity;
                this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

                if (activity.Conversation.ConversationType == ConversationTypes.Personal)
                {
                    if (activity.MembersAdded != null && activity.MembersAdded.Any(member => member.Id == activity.Recipient.Id))
                    {
                        await this.SendWelcomeCardInPersonalScopeAsync(turnContext);
                    }
                }
                else if (activity.Conversation.ConversationType == ConversationTypes.Channel)
                {
                    if (activity.MembersAdded != null && activity.MembersAdded.Any(member => member.Id == activity.Recipient.Id))
                    {
                        // If bot added to team, add team tab configuration with service URL.
                        await this.SendWelcomeCardInChannelAsync(turnContext);

                        var teamsDetails = activity.TeamsGetTeamInfo();

                        if (teamsDetails != null)
                        {
                            var teamTagConfiguration = new TeamTagEntity
                            {
                                CreatedDate = DateTime.UtcNow,
                                ServiceUrl = activity.ServiceUrl,
                                Tags = string.Empty,
                                TeamId = teamsDetails.Id,
                                CreatedByName = activity.From.Name,
                                UserAadId = activity.From.AadObjectId,
                            };

                            await this.teamTagStorageProvider.UpsertTeamTagAsync(teamTagConfiguration);
                        }
                    }
                    else if (activity.MembersRemoved != null && activity.MembersRemoved.Any(member => member.Id == activity.Recipient.Id))
                    {
                        // If bot removed from team, delete configured tags and digest preference settings.
                        var teamsDetails = activity.TeamsGetTeamInfo();
                        if (teamsDetails != null)
                        {
                            await this.teamTagStorageProvider.DeleteTeamTagAsync(teamsDetails.Id);
                            await this.teamPreferenceStorageProvider.DeleteTeamPreferenceAsync(teamsDetails.Id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while bot conversation update event.");
                throw;
            }
        }

        /// <summary>
        /// Invoked when the user opens the Messaging Extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains Messaging Extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionQuery query,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionQueryAsync), turnContext);

                var activity = turnContext.Activity;

                var messagingExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(activity.Value.ToString());
                var searchQuery = this.messagingExtensionHelper.GetSearchResult(messagingExtensionQuery);

                return new MessagingExtensionResponse
                {
                    ComposeExtension = await this.messagingExtensionHelper.GetTeamPostSearchResultAsync(searchQuery, messagingExtensionQuery.CommandId, activity.From.AadObjectId, messagingExtensionQuery.QueryOptions.Count, messagingExtensionQuery.QueryOptions.Skip),
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to handle the Messaging Extension command {turnContext.Activity.Name}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Invoked when task module fetch event is received from the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
#pragma warning disable CS1998 // Overriding method for task module fetch.
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
#pragma warning restore CS1998 // Overriding method for task module fetch.
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                taskModuleRequest = taskModuleRequest ?? throw new ArgumentNullException(nameof(taskModuleRequest));

                this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);

                var activity = turnContext.Activity;
                var postedValues = JsonConvert.DeserializeObject<BotCommand>(JObject.Parse(taskModuleRequest.Data.ToString()).SelectToken("data").ToString());
                var command = postedValues.Text;

                switch (command.ToUpperInvariant())
                {
                    case Constants.PreferenceSettings: // Preference settings command to set the tags in a team.
                        return new TaskModuleResponse
                        {
                            Task = new TaskModuleContinueResponse
                            {
                                Type = "continue",
                                Value = new TaskModuleTaskInfo()
                                {
                                    Url = $"{this.options.Value.AppBaseUri}/configurepreferences",
                                    Height = TaskModuleHeight,
                                    Width = TaskModuleWidth,
                                    Title = this.localizer.GetString("ApplicationName"),
                                },
                            },
                        };

                    default:
                        this.logger.LogInformation($"Received a command {command.ToUpperInvariant()} which is not supported.");
                        return default;
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching task module received by the bot.");
                throw;
            }
        }

        /// <summary>
        /// Invoked when a message activity is received from the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                var activity = turnContext.Activity;

                if (!string.IsNullOrEmpty(activity.Text))
                {
                    var command = activity.RemoveRecipientMention().Trim();

                    switch (command.ToUpperInvariant())
                    {
                        case Constants.HelpCommand: // Help command to get the information about the bot.
                            this.logger.LogInformation("Sending user help card.");
                            var userHelpCards = CarouselCard.GetUserHelpCards(this.options.Value.AppBaseUri);
                            await turnContext.SendActivityAsync(MessageFactory.Carousel(userHelpCards)).ConfigureAwait(false);
                            break;

                        case Constants.PreferenceSettings: // Preference command to get the card to setup the tags preference of a team.
                            await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetPreferenceCard(localizer: this.localizer)), cancellationToken).ConfigureAwait(false);
                            break;

                        default:
                            this.logger.LogInformation($"Received a command {command.ToUpperInvariant()} which is not supported.");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while message activity is received from the bot.");
                throw;
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                if (turnContext == null || taskModuleRequest == null)
                {
                    return new TaskModuleResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Type = "continue",
                            Value = new TaskModuleTaskInfo()
                            {
                                Url = $"{this.options.Value.AppBaseUri}/error",
                                Height = TaskModuleHeight,
                                Width = TaskModuleWidth,
                                Title = this.localizer.GetString("ApplicationName"),
                            },
                        },
                    };
                }

                var preferenceData = JsonConvert.DeserializeObject<Preference>(taskModuleRequest.Data?.ToString());

                if (preferenceData == null)
                {
                    this.logger.LogInformation($"Request data obtained on task module submit action is null.");
                    await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                    return null;
                }

                // If user clicks Cancel button in task module.
                if (preferenceData.Command == CloseCommand)
                {
                    return null;
                }

                if (preferenceData.Command == SubmitCommand)
                {
                    // Save or update digest preference for team.
                    if (preferenceData.ConfigureDetails != null)
                    {
                        var currentTeamPreferenceDetail = await this.teamPreferenceStorageProvider.GetTeamPreferenceAsync(preferenceData.ConfigureDetails.TeamId);
                        TeamPreferenceEntity teamPreferenceDetail;

                        if (currentTeamPreferenceDetail == null)
                        {
                            teamPreferenceDetail = new TeamPreferenceEntity
                            {
                                CreatedDate = DateTime.UtcNow,
                                DigestFrequency = preferenceData.ConfigureDetails.DigestFrequency,
                                Tags = preferenceData.ConfigureDetails.Tags,
                                TeamId = preferenceData.ConfigureDetails.TeamId,
                                UpdatedByName = turnContext.Activity.From.Name,
                                UpdatedByObjectId = turnContext.Activity.From.AadObjectId,
                                UpdatedDate = DateTime.UtcNow,
                                RowKey = preferenceData.ConfigureDetails.TeamId,
                            };
                        }
                        else
                        {
                            currentTeamPreferenceDetail.DigestFrequency = preferenceData.ConfigureDetails.DigestFrequency;
                            currentTeamPreferenceDetail.Tags = preferenceData.ConfigureDetails.Tags;
                            teamPreferenceDetail = currentTeamPreferenceDetail;
                        }

                        var upsertResult = await this.teamPreferenceStorageProvider.UpsertTeamPreferenceAsync(teamPreferenceDetail);
                    }
                    else
                    {
                        this.logger.LogInformation("Preference details received from task module is null.");
                        return new TaskModuleResponse
                        {
                            Task = new TaskModuleContinueResponse
                            {
                                Type = "continue",
                                Value = new TaskModuleTaskInfo()
                                {
                                    Url = $"{this.options.Value.AppBaseUri}/error",
                                    Height = TaskModuleHeight,
                                    Width = TaskModuleWidth,
                                    Title = this.localizer.GetString("ApplicationName"),
                                },
                            },
                        };
                    }
                }

                return null;
            }
#pragma warning disable CA1031 // Catching general exception for any errors occurred during saving data to table storage.
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exception for any errors occurred during saving data to table storage.
            {
                this.logger.LogError(ex, "Error in submit action of task module.");
                return new TaskModuleResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Type = "continue",
                        Value = new TaskModuleTaskInfo()
                        {
                            Url = $"{this.options.Value.AppBaseUri}/error",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ApplicationName"),
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Records event data to Application Insights telemetry client
        /// </summary>
        /// <param name="eventName">Name of the event.</param>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            var teamsChannelData = turnContext.Activity.GetChannelData<TeamsChannelData>();

            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
                { "teamId", teamsChannelData?.Team?.Id },
                { "channelId", teamsChannelData?.Channel?.Id },
            });
        }

        /// <summary>
        /// Sent welcome card to personal chat.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        /// <returns>A task that represents a response.</returns>
        private async Task SendWelcomeCardInPersonalScopeAsync(ITurnContext<IConversationUpdateActivity> turnContext)
        {
            this.logger.LogInformation($"Bot added in personal {turnContext.Activity.Conversation.Id}");
            var userStateAccessors = this.userState.CreateProperty<UserConversationState>(nameof(UserConversationState));
            var userConversationState = await userStateAccessors.GetAsync(turnContext, () => new UserConversationState());

            if (userConversationState?.IsWelcomeCardSent == null || userConversationState?.IsWelcomeCardSent == false)
            {
                userConversationState.IsWelcomeCardSent = true;
                await userStateAccessors.SetAsync(turnContext, userConversationState);

                var userWelcomeCardAttachment = WelcomeCard.GetWelcomeCardAttachmentForPersonal(
                    this.options.Value.AppBaseUri,
                    localizer: this.localizer,
                    this.botOptions.Value.ManifestId,
                    this.options.Value.DiscoverTabEntityId);

                await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment));
            }
        }

        /// <summary>
        /// Add user membership to storage if bot is installed in Team scope.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        /// <returns>A task that represents a response.</returns>
        private async Task SendWelcomeCardInChannelAsync(ITurnContext<IConversationUpdateActivity> turnContext)
        {
            this.logger.LogInformation($"Bot added in team {turnContext.Activity.Conversation.Id}");
            var userWelcomeCardAttachment = WelcomeCard.GetWelcomeCardAttachmentForTeam(this.options.Value.AppBaseUri, this.localizer);
            await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment));
        }
    }
}