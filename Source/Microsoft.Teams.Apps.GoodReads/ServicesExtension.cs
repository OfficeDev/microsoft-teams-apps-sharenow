// <copyright file="ServicesExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads
{
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Localization;
    using Microsoft.Azure.Search;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Azure;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Teams.Apps.GoodReads.Bot;
    using Microsoft.Teams.Apps.GoodReads.Common.BackgroundService;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Common.Providers;
    using Microsoft.Teams.Apps.GoodReads.Common.SearchServices;
    using Microsoft.Teams.Apps.GoodReads.Helpers;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.Teams.Apps.GoodReads.Models.Configuration;

    /// <summary>
    /// Class which helps to extend ServiceCollection.
    /// </summary>
    public static class ServicesExtension
    {
        /// <summary>
        /// Azure Search service index name for team post.
        /// </summary>
        private const string TeamPostIndexName = "team-post-index";

        /// <summary>
        /// Adds application configuration settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddConfigurationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            string appBaseUrl = configuration.GetValue<string>("App:AppBaseUri");
            string discoverTabEntityId = configuration.GetValue<string>("DiscoverTabEntityId");

            services.Configure<GoodReadsActivityHandlerOptions>(options =>
            {
                options.AppBaseUri = appBaseUrl;
                options.DiscoverTabEntityId = discoverTabEntityId;
            });

            services.Configure<BotSettings>(options =>
            {
                options.SecurityKey = configuration.GetValue<string>("App:SecurityKey");
                options.AppBaseUri = appBaseUrl;
                options.TenantId = configuration.GetValue<string>("App:TenantId");
                options.MedianFirstRetryDelay = configuration.GetValue<double>("RetryPolicy:medianFirstRetryDelay");
                options.RetryCount = configuration.GetValue<int>("RetryPolicy:retryCount");
                options.ManifestId = configuration.GetValue<string>("App:ManifestId");
                options.MicrosoftAppId = configuration.GetValue<string>("MicrosoftAppId");
                options.MicrosoftAppPassword = configuration.GetValue<string>("MicrosoftAppPassword");
            });

            services.Configure<AzureActiveDirectorySettings>(options =>
            {
                options.TenantId = configuration.GetValue<string>("AzureAd:TenantId");
                options.ClientId = configuration.GetValue<string>("AzureAd:ClientId");
                options.ApplicationIdURI = configuration.GetValue<string>("AzureAd:ApplicationIdURI");
                options.ValidIssuers = configuration.GetValue<string>("AzureAd:ValidIssuers");
                options.Instance = configuration.GetValue<string>("AzureAd:Instance");
            });

            services.Configure<TelemetrySetting>(options =>
            {
                options.InstrumentationKey = configuration.GetValue<string>("ApplicationInsights:InstrumentationKey");
            });

            services.Configure<StorageSettings>(options =>
            {
                options.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });

            services.Configure<SearchServiceSetting>(searchServiceSettings =>
            {
                searchServiceSettings.SearchServiceName = configuration.GetValue<string>("SearchService:SearchServiceName");
                searchServiceSettings.SearchServiceQueryApiKey = configuration.GetValue<string>("SearchService:SearchServiceQueryApiKey");
                searchServiceSettings.SearchServiceAdminApiKey = configuration.GetValue<string>("SearchService:SearchServiceAdminApiKey");
                searchServiceSettings.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });
        }

        /// <summary>
        /// Adds helpers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddHelpers(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddApplicationInsightsTelemetry(configuration.GetValue<string>("ApplicationInsights:InstrumentationKey"));

            services.AddSingleton<IPostStorageProvider, PostStorageProvider>();
            services.AddSingleton<ITeamPreferenceStorageProvider, TeamPreferenceStorageProvider>();
            services.AddSingleton<IUserPrivatePostStorageProvider, UserPrivatePostStorageProvider>();
            services.AddSingleton<IUserVoteStorageProvider, UserVoteStorageProvider>();
            services.AddSingleton<ITeamTagStorageProvider, TeamTagStorageProvider>();
            services.AddSingleton<ITeamTagStorageProvider, TeamTagStorageProvider>();
            services.AddSingleton<PostTypeHelper>();
            services.AddSingleton<IPostSearchService, PostSearchService>();
            services.AddSingleton<IMessagingExtensionHelper, MessagingExtensionHelper>();
            services.AddSingleton<IPostStorageHelper, PostStorageHelper>();
            services.AddSingleton<ITeamPreferenceStorageHelper, TeamPreferenceStorageHelper>();
#pragma warning disable CA2000 // This is singleton which has lifetime same as the app
            services.AddSingleton<ISearchServiceClient>(new SearchServiceClient(configuration.GetValue<string>("SearchService:SearchServiceName"), new SearchCredentials(configuration.GetValue<string>("SearchService:SearchServiceAdminApiKey"))));
            services.AddSingleton<ISearchIndexClient>(new SearchIndexClient(configuration.GetValue<string>("SearchService:SearchServiceName"), TeamPostIndexName, new SearchCredentials(configuration.GetValue<string>("SearchService:SearchServiceQueryApiKey"))));
#pragma warning restore CA2000 // This is singleton which has lifetime same as the app
            services.AddHostedService<DigestNotificationBackgroundService>();
            services.AddSingleton<IDigestNotificationHelper, DigestNotificationHelper>();
        }

        /// <summary>
        /// Adds user state and conversation state to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddBotStates(this IServiceCollection services, IConfiguration configuration)
        {
            // Create the User state. (Used in this bot's Dialog implementation.)
            services.AddSingleton<UserState>();

            // Create the Conversation state. (Used by the Dialog system itself.)
            services.AddSingleton<ConversationState>();

            // For conversation state.
            services.AddSingleton<IStorage>(new AzureBlobStorage(configuration.GetValue<string>("Storage:ConnectionString"), "bot-state"));
        }

        /// <summary>
        /// Adds localization.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddLocalization(this IServiceCollection services, IConfiguration configuration)
        {
            // Add i18n.
            services.AddLocalization(options => options.ResourcesPath = "Resources");

            services.Configure<RequestLocalizationOptions>(options =>
            {
                var defaultCulture = CultureInfo.GetCultureInfo(configuration.GetValue<string>("i18n:DefaultCulture"));
                var supportedCultures = configuration.GetValue<string>("i18n:SupportedCultures").Split(',')
                    .Select(culture => CultureInfo.GetCultureInfo(culture))
                    .ToList();

                options.DefaultRequestCulture = new RequestCulture(defaultCulture);
                options.SupportedCultures = supportedCultures;
                options.SupportedUICultures = supportedCultures;

                options.RequestCultureProviders = new List<IRequestCultureProvider>
                {
                    new GoodReadsLocalizationCultureProvider(),
                };
            });
        }

        /// <summary>
        /// Adds credential providers for authentication.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCredentialProviders(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services.AddSingleton(new MicrosoftAppCredentials(configuration.GetValue<string>("MicrosoftAppId"), configuration.GetValue<string>("MicrosoftAppPassword")));
        }
    }
}