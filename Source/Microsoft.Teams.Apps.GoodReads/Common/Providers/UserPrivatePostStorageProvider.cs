// <copyright file="UserPrivatePostStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Providers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.Teams.Apps.GoodReads.Models.Configuration;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which help to create, get, update or delete team post data in user's private list.
    /// </summary>
    public class UserPrivatePostStorageProvider : BaseStorageProvider, IUserPrivatePostStorageProvider
    {
        /// <summary>
        /// Represents user's private post entity name.
        /// </summary>
        private const string UserPrivatePostEntityName = "UserPrivatePostEntity";

        /// <summary>
        /// Initializes a new instance of the <see cref="UserPrivatePostStorageProvider"/> class.
        /// Handles storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public UserPrivatePostStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<BaseStorageProvider> logger)
            : base(options?.Value.ConnectionString, UserPrivatePostEntityName, logger)
        {
        }

        /// <summary>
        /// Get user's private list of posts data.
        /// </summary>
        /// <param name="userId">User id for which need to fetch data.</param>
        /// <returns>A task that represent collection to hold user's private list of posts data.</returns>
        public async Task<IEnumerable<string>> GetUserPrivatePostsIdAsync(string userId)
        {
            await this.EnsureInitializedAsync();
            var partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, userId);

            TableQuery<UserPrivatePostEntity> query = new TableQuery<UserPrivatePostEntity>().Where(partitionKeyCondition);
            TableContinuationToken continuationToken = null;
            var userPrivatePostCollection = new List<UserPrivatePostEntity>();

            do
            {
                var queryResult = await this.GoodReadsCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    userPrivatePostCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return userPrivatePostCollection.OrderByDescending(post => post.CreatedDate).Select(privatePost => privatePost.PostId);
        }

        /// <summary>
        /// Delete private post from user's private list.
        /// </summary>
        /// <param name="postId">Holds private post id.</param>
        /// <param name="userId">Azure Active Directory id of user.</param>
        /// <returns>A boolean that represents private post is successfully deleted or not.</returns>
        public async Task<bool> DeleteUserPrivatePostAsync(string postId, string userId)
        {
            await this.EnsureInitializedAsync();

            var retrieveOperation = TableOperation.Retrieve<UserPrivatePostEntity>(userId, postId);
            var queryResult = await this.GoodReadsCloudTable.ExecuteAsync(retrieveOperation);

            if (queryResult?.Result != null)
            {
                TableOperation deleteOperation = TableOperation.Delete((UserPrivatePostEntity)queryResult.Result);
                var result = await this.GoodReadsCloudTable.ExecuteAsync(deleteOperation);

                return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
            }

            return false;
        }

        /// <summary>
        /// Stores or update post data in user's private list.
        /// </summary>
        /// <param name="entity">Holds user post detail.</param>
        /// <returns>A boolean that represents user private post is successfully saved/updated or not.</returns>
        public async Task<bool> UpsertUserPrivatPostAsync(UserPrivatePostEntity entity)
        {
            var result = await this.StoreOrUpdatePrivatePostAsync(entity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Stores or update post data in user's private list.
        /// </summary>
        /// <param name="entity">Represents user private post entity object.</param>
        /// <returns>A task that represents user private post is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdatePrivatePostAsync(UserPrivatePostEntity entity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(entity);
            return await this.GoodReadsCloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}
