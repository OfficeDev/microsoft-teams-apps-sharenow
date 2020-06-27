// <copyright file="UserVoteStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.Teams.Apps.GoodReads.Models.Configuration;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps to create, get, update or delete user vote data.
    /// </summary>
    public class UserVoteStorageProvider : BaseStorageProvider, IUserVoteStorageProvider
    {
        /// <summary>
        /// Represents user vote entity name.
        /// </summary>
        private const string UserVoteEntityName = "UserVoteEntity";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserVoteStorageProvider"/> class.
        /// Handles storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public UserVoteStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<BaseStorageProvider> logger)
            : base(options?.Value.ConnectionString, UserVoteEntityName, logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Get all user votes.
        /// </summary>
        /// <param name="userId">Represent Azure Active Directory id of user.</param>
        /// <returns>A task that represents a collection of user votes.</returns>
        public async Task<IEnumerable<UserVoteEntity>> GetUserVotesAsync(string userId)
        {
            await this.EnsureInitializedAsync();

            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, userId);

            List<UserVoteEntity> userVotes = new List<UserVoteEntity>();
            TableContinuationToken continuationToken = null;
            TableQuery<UserVoteEntity> query = new TableQuery<UserVoteEntity>().Where(partitionKeyCondition);

            do
            {
                var queryResult = await this.GoodReadsCloudTable.ExecuteQuerySegmentedAsync(query, null);
                if (queryResult?.Results != null)
                {
                    userVotes.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return userVotes;
        }

        /// <summary>
        /// Get user vote for post.
        /// </summary>
        /// <param name="userId">Represent Azure Active Directory id of user.</param>
        /// <param name="postId">Post Id for which user has voted.</param>
        /// <returns>A task that represents a collection of user votes.</returns>
        public async Task<UserVoteEntity> GetUserVoteForPostAsync(string userId, string postId)
        {
            await this.EnsureInitializedAsync();

            var retrieveOperation = TableOperation.Retrieve<UserVoteEntity>(userId, postId);
            var queryResult = await this.GoodReadsCloudTable.ExecuteAsync(retrieveOperation);

            if (queryResult?.Result != null)
            {
                return (UserVoteEntity)queryResult.Result;
            }

            return null;
        }

        /// <summary>
        /// Stores or update user votes data.
        /// </summary>
        /// <param name="voteEntity">Holds user vote entity data.</param>
        /// <returns>A boolean that represents user vote entity is successfully saved/updated or not.</returns>
        public async Task<bool> UpsertUserVoteAsync(UserVoteEntity voteEntity)
        {
            var result = await this.StoreOrUpdateUserVoteAsync(voteEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Delete user vote data.
        /// </summary>
        /// <param name="postId">Represents post id.</param>
        /// <param name="userId">Represent Azure Active Directory id of user.</param>
        /// <returns>A boolean that represents user vote data is successfully deleted or not.</returns>
        public async Task<bool> DeleteUserVoteAsync(string postId, string userId)
        {
            try
            {
                await this.EnsureInitializedAsync();

                var retrieveOperation = TableOperation.Retrieve<UserVoteEntity>(userId, postId);
                var queryResult = await this.GoodReadsCloudTable.ExecuteAsync(retrieveOperation);

                if (queryResult?.Result != null)
                {
                    TableOperation deleteOperation = TableOperation.Delete((ITableEntity)queryResult.Result);
                    var result = await this.GoodReadsCloudTable.ExecuteAsync(deleteOperation);

                    return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
                }
            }
#pragma warning disable CA1031 // Catching generic exceptions to log error in telemetry and caller to get only operation success/failure status
            catch (Exception ex)
#pragma warning restore CA1031 // Catching generic exceptions to log error in telemetry and caller to get only operation success/failure status
            {
                this.logger.LogError(ex, "Exception occurred while performing delete user vote operation.");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Stores or update user votes data.
        /// </summary>
        /// <param name="voteEntity">Holds user vote entity data.</param>
        /// <returns>A task that represents user vote entity data is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateUserVoteAsync(UserVoteEntity voteEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(voteEntity);
            return await this.GoodReadsCloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}
