// <copyright file="TeamPreferenceStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Providers
{
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
    /// Implements storage provider which helps to create, get or update team preferences data.
    /// </summary>
    public class TeamPreferenceStorageProvider : BaseStorageProvider, ITeamPreferenceStorageProvider
    {
        /// <summary>
        /// Represents team preference entity name.
        /// </summary>
        private const string TeamPreferenceEntityName = "TeamPreferenceEntity";

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamPreferenceStorageProvider"/> class.
        /// Handles storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public TeamPreferenceStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<BaseStorageProvider> logger)
            : base(options?.Value.ConnectionString, TeamPreferenceEntityName, logger)
        {
        }

        /// <summary>
        /// Get team preference data.
        /// </summary>
        /// <param name="teamId">Team Id for which need to fetch data.</param>
        /// <returns>A task that represents an object to hold team preference data.</returns>
        public async Task<TeamPreferenceEntity> GetTeamPreferenceAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            var retrieveOperation = TableOperation.Retrieve<TeamPreferenceEntity>(teamId, teamId);
            var queryResult = await this.GoodReadsCloudTable.ExecuteAsync(retrieveOperation);
            if (queryResult?.Result != null)
            {
                return (TeamPreferenceEntity)queryResult?.Result;
            }

            return null;
        }

        /// <summary>
        /// Get team preferences data.
        /// </summary>
        /// <param name="digestFrequency">Digest frequency text for notification like Monthly/Weekly.</param>
        /// <returns>A task that represent collection to hold team preferences data.</returns>
        public async Task<IEnumerable<TeamPreferenceEntity>> GetTeamPreferencesByDigestFrequencyAsync(string digestFrequency)
        {
            await this.EnsureInitializedAsync();

            var digestFrequencyCondition = TableQuery.GenerateFilterCondition(nameof(TeamPreferenceEntity.DigestFrequency), QueryComparisons.Equal, digestFrequency);

            TableQuery<TeamPreferenceEntity> query = new TableQuery<TeamPreferenceEntity>().Where(digestFrequencyCondition);
            TableContinuationToken continuationToken = null;
            var teamPreferenceCollection = new List<TeamPreferenceEntity>();

            do
            {
                var queryResult = await this.GoodReadsCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);

                if (queryResult?.Results != null)
                {
                    teamPreferenceCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return teamPreferenceCollection;
        }

        /// <summary>
        /// Delete team preference if Bot is uninstalled.
        /// </summary>
        /// <param name="teamId">Holds team id.</param>
        /// <returns>A boolean that represents team tags is successfully deleted or not.</returns>
        public async Task<bool> DeleteTeamPreferenceAsync(string teamId)
        {
            await this.EnsureInitializedAsync();

            var retrieveOperation = TableOperation.Retrieve<TeamPreferenceEntity>(teamId, teamId);
            var queryResult = await this.GoodReadsCloudTable.ExecuteAsync(retrieveOperation);
            if (queryResult?.Result != null)
            {
                TableOperation deleteOperation = TableOperation.Delete((TeamPreferenceEntity)queryResult.Result);
                var result = await this.GoodReadsCloudTable.ExecuteAsync(deleteOperation);
                return result.HttpStatusCode == (int)HttpStatusCode.OK;
            }

            return false;
        }

        /// <summary>
        /// Stores or update team preference data.
        /// </summary>
        /// <param name="teamPreferenceEntity">Represents team preference entity object.</param>
        /// <returns>A boolean that represents team preference entity is successfully saved/updated or not.</returns>
        public async Task<bool> UpsertTeamPreferenceAsync(TeamPreferenceEntity teamPreferenceEntity)
        {
            var result = await this.StoreOrUpdateTeamPreferenceAsync(teamPreferenceEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Stores or update team preference data in storage.
        /// </summary>
        /// <param name="teamPreferenceEntity">Holds team preference detail entity data.</param>
        /// <returns>A task that represents team preference entity data is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateTeamPreferenceAsync(TeamPreferenceEntity teamPreferenceEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(teamPreferenceEntity);
            return await this.GoodReadsCloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}
