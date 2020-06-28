// <copyright file="PostSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.SearchServices
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Rest.Azure;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.Teams.Apps.GoodReads.Models.Configuration;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Post Search service which helps in creating index, indexer and data source if it doesn't exist
    /// for indexing table which will be used for search by Messaging Extension.
    /// </summary>
    public class PostSearchService : IPostSearchService, IDisposable
    {
        /// <summary>
        /// Azure Search service index name.
        /// </summary>
        private const string IndexName = "team-post-index";

        /// <summary>
        /// Azure Search service indexer name.
        /// </summary>
        private const string IndexerName = "team-post-indexer";

        /// <summary>
        /// Azure Search service data source name.
        /// </summary>
        private const string DataSourceName = "team-post-storage";

        /// <summary>
        /// Table name where team post data will get saved.
        /// </summary>
        private const string TeamPostTableName = "TeamPostEntity";

        /// <summary>
        /// Represents the sorting type as popularity means to sort the data based on number of votes.
        /// </summary>
        private const int SortByPopular = 1;

        /// <summary>
        /// Azure Search service maximum search result count for team post entity.
        /// </summary>
        private const int ApiSearchResultCount = 1500;

        /// <summary>
        /// Retry policy with jitter.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private readonly AsyncRetryPolicy retryPolicy;

        /// <summary>
        /// Used to initialize task.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Instance of Azure Search service client.
        /// </summary>
        private readonly ISearchServiceClient searchServiceClient;

        /// <summary>
        /// Instance of Azure Search index client.
        /// </summary>
        private readonly ISearchIndexClient searchIndexClient;

        /// <summary>
        /// Instance of post storage helper to update post and get information of posts.
        /// </summary>
        private readonly IPostStorageProvider postStorageProvider;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<PostSearchService> logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly SearchServiceSetting options;

        /// <summary>
        /// Flag: Has Dispose already been called?
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="PostSearchService"/> class.
        /// </summary>
        /// <param name="optionsAccessor">A set of key/value application configuration properties.</param>
        /// <param name="postStorageProvider">Post storage provider dependency injection.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="searchServiceClient">Search service client dependency injection.</param>
        /// <param name="searchIndexClient">Search index client dependency injection.</param>
        public PostSearchService(
            IOptions<SearchServiceSetting> optionsAccessor,
            IPostStorageProvider postStorageProvider,
            ILogger<PostSearchService> logger,
            ISearchServiceClient searchServiceClient,
            ISearchIndexClient searchIndexClient)
        {
            optionsAccessor = optionsAccessor ?? throw new ArgumentNullException(nameof(optionsAccessor));

            this.options = optionsAccessor.Value;
            var searchServiceValue = this.options.SearchServiceName;
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync());
            this.postStorageProvider = postStorageProvider;
            this.logger = logger;
            this.searchServiceClient = searchServiceClient;
            this.searchIndexClient = searchIndexClient;
            this.retryPolicy = Policy.Handle<CloudException>(
                ex => (int)ex.Response.StatusCode == StatusCodes.Status409Conflict ||
                (int)ex.Response.StatusCode == StatusCodes.Status429TooManyRequests)
                .WaitAndRetryAsync(Backoff.LinearBackoff(TimeSpan.FromMilliseconds(2000), 2));
        }

        /// <summary>
        /// Provide search result for table to be used by user's based on Azure Search service.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="searchQuery">Query which the user had typed in Messaging Extension search field.</param>
        /// <param name="userObjectId">Azure Active Directory object id of the user.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="filterQuery">Filter bar based query.</param>
        /// <returns>List of search results.</returns>
        public async Task<IEnumerable<PostEntity>> GetPostsAsync(
            PostSearchScope searchScope,
            string searchQuery,
            string userObjectId,
            int? count = null,
            int? skip = null,
            int? sortBy = null,
            string filterQuery = null)
        {
            await this.EnsureInitializedAsync();
            var searchParameters = this.InitializeSearchParameters(searchScope, userObjectId, count, skip, sortBy, filterQuery);

            SearchContinuationToken continuationToken = null;
            var posts = new List<PostEntity>();
            var postSearchResult = await this.searchIndexClient.Documents.SearchAsync<PostEntity>(searchQuery, searchParameters);

            if (postSearchResult?.Results != null)
            {
                posts.AddRange(postSearchResult.Results.Select(p => p.Document));
                continuationToken = postSearchResult.ContinuationToken;
            }

            if (continuationToken == null)
            {
                return posts;
            }

            do
            {
                var searchResult = await this.searchIndexClient.Documents.ContinueSearchAsync<PostEntity>(continuationToken);

                if (searchResult?.Results != null)
                {
                    posts.AddRange(searchResult.Results.Select(p => p.Document));
                    continuationToken = searchResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return posts;
        }

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task RecreateSearchServiceIndexAsync()
        {
            try
            {
                await this.CreateSearchIndexAsync();
                await this.CreateDataSourceAsync();
                await this.CreateIndexerAsync();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Run the indexer on demand.
        /// </summary>
        /// <returns>A task that represents the work queued to execute</returns>
        public async Task RunIndexerOnDemandAsync()
        {
            // Retry once after 1 second if conflict occurs during indexer run.
            // If conflict occurs again means another index run is in progress and it will index data for which first failure occurred.
            // Hence ignore second conflict and continue.
            var requestId = Guid.NewGuid().ToString();

            try
            {
                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        this.logger.LogInformation($"On-demand indexer run request #{requestId} - start");
                        await this.searchServiceClient.Indexers.RunAsync(IndexerName);
                        this.logger.LogInformation($"On-demand indexer run request #{requestId} - complete");
                    }
                    catch (CloudException ex)
                    {
                        this.logger.LogError(ex, $"Failed to run on-demand indexer run for request #{requestId}: {ex.Message}");
                        throw;
                    }
                });
            }
            catch (CloudException ex)
            {
                this.logger.LogError(ex, $"Failed to run on-demand indexer for retry. Request #{requestId}: {ex.Message}");
            }
        }

        /// <summary>
        /// Dispose search service instance.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.searchServiceClient.Dispose();
                this.searchIndexClient.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Create index, indexer and data source if doesn't exist.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync()
        {
            try
            {
                // When there is no post created by user and Messaging Extension is open, table initialization is required here before creating search index or data source or indexer.
                await this.postStorageProvider.GetPostAsync(postCreatedByuserId: string.Empty, postId: string.Empty);
                await this.RecreateSearchServiceIndexAsync();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to initialize Azure Search Service: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Create index in Azure Search service if it doesn't exist.
        /// </summary>
        /// <returns><see cref="Task"/> That represents index is created if it is not created.</returns>
        private async Task CreateSearchIndexAsync()
        {
            if (await this.searchServiceClient.Indexes.ExistsAsync(IndexName))
            {
                await this.searchServiceClient.Indexes.DeleteAsync(IndexName);
            }

            var tableIndex = new Index()
            {
                Name = IndexName,
                Fields = FieldBuilder.BuildForType<PostEntity>(),
            };
            await this.searchServiceClient.Indexes.CreateAsync(tableIndex);
        }

        /// <summary>
        /// Create data source if it doesn't exist in Azure Search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents data source is added to Azure Search service.</returns>
        private async Task CreateDataSourceAsync()
        {
            if (await this.searchServiceClient.DataSources.ExistsAsync(DataSourceName))
            {
                return;
            }

            var dataSource = DataSource.AzureTableStorage(
                DataSourceName,
                this.options.ConnectionString,
                TeamPostTableName,
                query: null,
                new SoftDeleteColumnDeletionDetectionPolicy("IsRemoved", true));

            await this.searchServiceClient.DataSources.CreateAsync(dataSource);
        }

        /// <summary>
        /// Create indexer if it doesn't exist in Azure Search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents indexer is created if not available in Azure Search service.</returns>
        private async Task CreateIndexerAsync()
        {
            if (await this.searchServiceClient.Indexers.ExistsAsync(IndexerName))
            {
                await this.searchServiceClient.Indexers.DeleteAsync(IndexerName);
            }

            var indexer = new Indexer()
            {
                Name = IndexerName,
                DataSourceName = DataSourceName,
                TargetIndexName = IndexName,
            };

            await this.searchServiceClient.Indexers.CreateAsync(indexer);
            await this.searchServiceClient.Indexers.RunAsync(IndexerName);
        }

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        private Task EnsureInitializedAsync()
        {
            return this.initializeTask.Value;
        }

        /// <summary>
        /// Initialization of search service parameters which will help in searching the documents.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="userObjectId">Azure Active Directory object id of the user.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="filterQuery">Filter bar based query.</param>
        /// <returns>Represents an search parameter object.</returns>
        private SearchParameters InitializeSearchParameters(
            PostSearchScope searchScope,
            string userObjectId,
            int? count = null,
            int? skip = null,
            int? sortBy = null,
            string filterQuery = null)
        {
            SearchParameters searchParameters = new SearchParameters()
            {
                Top = count ?? ApiSearchResultCount,
                Skip = skip ?? 0,
                IncludeTotalResultCount = false,
                Select = new[]
                {
                    nameof(PostEntity.PostId),
                    nameof(PostEntity.Type),
                    nameof(PostEntity.Title),
                    nameof(PostEntity.Description),
                    nameof(PostEntity.ContentUrl),
                    nameof(PostEntity.Tags),
                    nameof(PostEntity.CreatedDate),
                    nameof(PostEntity.CreatedByName),
                    nameof(PostEntity.UpdatedDate),
                    nameof(PostEntity.UserId),
                    nameof(PostEntity.TotalVotes),
                    nameof(PostEntity.IsRemoved),
                },
                SearchFields = new[] { nameof(PostEntity.Title) },
                Filter = !string.IsNullOrEmpty(filterQuery) ? filterQuery : string.Empty,
            };

            switch (searchScope)
            {
                case PostSearchScope.AllItems:
                    searchParameters.OrderBy = new[] { $"{nameof(PostEntity.UpdatedDate)} desc" };
                    break;

                case PostSearchScope.PostedByMe:
                    searchParameters.Filter = $"{nameof(PostEntity.UserId)} eq '{userObjectId}' ";
                    searchParameters.OrderBy = new[] { $"{nameof(PostEntity.UpdatedDate)} desc" };
                    break;

                case PostSearchScope.Popular:
                    searchParameters.OrderBy = new[] { $"{nameof(PostEntity.TotalVotes)} desc" };
                    break;

                case PostSearchScope.TeamPreferenceTags:
                    searchParameters.SearchFields = new[] { nameof(PostEntity.Tags) };
                    searchParameters.Top = 5000;
                    searchParameters.Select = new[] { nameof(PostEntity.Tags) };
                    break;

                case PostSearchScope.FilterAsPerTeamTags:
                    searchParameters.OrderBy = new[] { $"{nameof(PostEntity.UpdatedDate)} desc" };
                    searchParameters.SearchFields = new[] { nameof(PostEntity.Tags) };
                    break;

                case PostSearchScope.FilterPostsAsPerDateRange:
                    searchParameters.OrderBy = new[] { $"{nameof(PostEntity.UpdatedDate)} desc" };
                    searchParameters.Top = 200;
                    break;

                case PostSearchScope.UniqueUserNames:
                    searchParameters.OrderBy = new[] { $"{nameof(PostEntity.UpdatedDate)} desc" };
                    searchParameters.Select = new[] { nameof(PostEntity.CreatedByName), nameof(PostEntity.UserId) };
                    break;

                case PostSearchScope.SearchTeamPostsForTitleText:
                    searchParameters.OrderBy = new[] { $"{nameof(PostEntity.UpdatedDate)} desc" };
                    searchParameters.QueryType = QueryType.Full;
                    searchParameters.SearchFields = new[] { nameof(PostEntity.Title) };
                    break;

                case PostSearchScope.FilterTeamPosts:
                    if (sortBy != null)
                    {
                        searchParameters.OrderBy = sortBy == SortByPopular ? new[] { $"{nameof(PostEntity.TotalVotes)} desc" } : new[] { $"{nameof(PostEntity.UpdatedDate)} desc" };
                    }

                    searchParameters.SearchFields = new[] { nameof(PostEntity.Tags) };
                    break;
            }

            return searchParameters;
        }
    }
}