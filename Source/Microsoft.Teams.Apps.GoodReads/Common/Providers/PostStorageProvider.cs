// <copyright file="PostStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Providers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.Teams.Apps.GoodReads.Models.Configuration;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps to create, get, update or delete posts data.
    /// </summary>
    public class PostStorageProvider : BaseStorageProvider, IPostStorageProvider
    {
        /// <summary>
        /// Represents post entity name.
        /// </summary>
        private const string PostEntityName = "TeamPostEntity";

        /// <summary>
        /// Represent a column name.
        /// </summary>
        private const string IsRemovedColumnName = "IsRemoved";

        /// <summary>
        /// Initializes a new instance of the <see cref="PostStorageProvider"/> class.
        /// Handles storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public PostStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<BaseStorageProvider> logger)
            : base(options?.Value.ConnectionString, PostEntityName, logger)
        {
        }

        /// <summary>
        /// Get posts data.
        /// </summary>
        /// <param name="isRemoved">Represent a post is deleted or not.</param>
        /// <returns>A task that represent collection to hold posts.</returns>
        public async Task<IEnumerable<PostEntity>> GetPostsAsync(bool isRemoved)
        {
            await this.EnsureInitializedAsync();

            string isRemovedCondition = TableQuery.GenerateFilterConditionForBool(IsRemovedColumnName, QueryComparisons.Equal, isRemoved);
            TableQuery<PostEntity> query = new TableQuery<PostEntity>().Where(isRemovedCondition);
            TableContinuationToken continuationToken = null;
            var postCollection = new List<PostEntity>();

            do
            {
                var queryResult = await this.GoodReadsCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    postCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return postCollection;
        }

        /// <summary>
        /// Get post data.
        /// </summary>
        /// <param name="postCreatedByuserId">User id to fetch the post details.</param>
        /// <param name="postId">Post id to fetch the post details.</param>
        /// <returns>A task that represent a object to hold post data.</returns>
        public async Task<PostEntity> GetPostAsync(string postCreatedByuserId, string postId)
        {
            // When there is no post created by user and Messaging Extension is open, table initialization is required here before creating search index or data source or indexer.
            await this.EnsureInitializedAsync();

            if (string.IsNullOrEmpty(postId) || string.IsNullOrEmpty(postCreatedByuserId))
            {
                return null;
            }

            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, postCreatedByuserId);
            string postIdCondition = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, postId);
            var combinedPartitionFilter = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, postIdCondition);

            string isRemovedCondition = TableQuery.GenerateFilterConditionForBool(IsRemovedColumnName, QueryComparisons.Equal, false);
            var combinedFilter = TableQuery.CombineFilters(combinedPartitionFilter, TableOperators.And, isRemovedCondition);

            TableQuery<PostEntity> query = new TableQuery<PostEntity>().Where(combinedFilter);
            var queryResult = await this.GoodReadsCloudTable.ExecuteQuerySegmentedAsync(query, null);

            return queryResult?.FirstOrDefault();
        }

        /// <summary>
        /// Stores or update post details data.
        /// </summary>
        /// <param name="postEntity">Holds post detail entity data.</param>
        /// <returns>A boolean that represents post entity data is successfully saved/updated or not.</returns>
        public async Task<bool> UpsertPostAsync(PostEntity postEntity)
        {
            var result = await this.StoreOrUpdatePostAsync(postEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get posts as per the user's private list of post.
        /// </summary>
        /// <param name="postIds">A collection of user private post id's.</param>
        /// <returns>A task that represent collection to hold posts data.</returns>
        public async Task<IEnumerable<PostEntity>> GetFilteredUserPrivatePostsAsync(IEnumerable<string> postIds)
        {
            await this.EnsureInitializedAsync();
            string privatePostCondition = this.CreateUserPrivatePostsFilter(postIds);
            string isRemovedCondition = TableQuery.GenerateFilterConditionForBool(IsRemovedColumnName, QueryComparisons.Equal, false);
            var combinedFilter = TableQuery.CombineFilters(privatePostCondition, TableOperators.And, isRemovedCondition);

            TableQuery<PostEntity> query = new TableQuery<PostEntity>().Where(combinedFilter);
            TableContinuationToken continuationToken = null;
            var postCollection = new List<PostEntity>();
            do
            {
                var queryResult = await this.GoodReadsCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    postCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return postCollection;
        }

        /// <summary>
        /// Get combined filter condition for user private posts data.
        /// </summary>
        /// <param name="postIds">List of user private posts id.</param>
        /// <returns>Returns combined filter for user private posts.</returns>
        private string CreateUserPrivatePostsFilter(IEnumerable<string> postIds)
        {
            var postIdConditions = new List<string>();
            StringBuilder combinedPostIdFilter = new StringBuilder();

            postIds = postIds.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();

            foreach (var postId in postIds)
            {
                postIdConditions.Add("(" + TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, postId) + ")");
            }

            if (postIdConditions.Count >= 2)
            {
                var posts = postIdConditions.Take(postIdConditions.Count - 1).ToList();

                posts.ForEach(postCondition =>
                {
                    combinedPostIdFilter.Append($"{postCondition} {"or"} ");
                });

                combinedPostIdFilter.Append($"{postIdConditions.Last()}");

                return combinedPostIdFilter.ToString();
            }
            else
            {
                return TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, postIds.FirstOrDefault());
            }
        }

        /// <summary>
        /// Stores or update post details data.
        /// </summary>
        /// <param name="entity">Holds post detail entity data.</param>
        /// <returns>A task that represents post entity data is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdatePostAsync(PostEntity entity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(entity);
            return await this.GoodReadsCloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}
