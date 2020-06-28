// <copyright file="BaseStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Providers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which initializes table if not exists and provide table client instance.
    /// </summary>
    public class BaseStorageProvider
    {
        /// <summary>
        /// Storage connection string.
        /// </summary>
        private readonly string connectionString;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseStorageProvider"/> class.
        /// Handles Microsoft Azure Table creation.
        /// </summary>
        /// <param name="connectionString">Azure Table storage connection string.</param>
        /// <param name="tableName">Azure Table storage table name.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public BaseStorageProvider(
            string connectionString,
            string tableName,
            ILogger<BaseStorageProvider> logger)
        {
            this.InitializeTask = new Lazy<Task>(() => this.InitializeAsync());
            this.connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
            this.TableName = tableName;
            this.logger = logger;
        }

        /// <summary>
        /// Gets or sets task for initialization.
        /// </summary>
        protected Lazy<Task> InitializeTask { get; set; }

        /// <summary>
        /// Gets or sets Storage table name.
        /// </summary>
        protected string TableName { get; set; }

        /// <summary>
        /// Gets or sets a table in the storage.
        /// </summary>
        protected CloudTable GoodReadsCloudTable { get; set; }

        /// <summary>
        /// Ensures storage should be created before working on table.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        protected async Task EnsureInitializedAsync()
        {
            await this.InitializeTask.Value;
        }

        /// <summary>
        /// Combine two filter conditions in a single filter string.
        /// </summary>
        /// <param name="teamIdFilter">First filter condition.</param>
        /// <param name="partitionKeyFilter">Second filter condition.</param>
        /// <returns> single filter string by combining two filter conditions.</returns>
        protected string CombineFilters(string teamIdFilter, string partitionKeyFilter)
        {
            if (string.IsNullOrWhiteSpace(teamIdFilter) && string.IsNullOrWhiteSpace(partitionKeyFilter))
            {
                return string.Empty;
            }
            else if (string.IsNullOrWhiteSpace(teamIdFilter))
            {
                return partitionKeyFilter;
            }
            else if (string.IsNullOrWhiteSpace(partitionKeyFilter))
            {
                return teamIdFilter;
            }

            return TableQuery.CombineFilters(teamIdFilter, TableOperators.And, partitionKeyFilter);
        }

        /// <summary>
        /// Create tables if it doesn't exist.
        /// </summary>
        /// <returns>Asynchronous task which represents table is created if its not existing.</returns>
        private async Task InitializeAsync()
        {
            try
            {
               CloudStorageAccount storageAccount = CloudStorageAccount.Parse(this.connectionString);
               CloudTableClient cloudTableClient = storageAccount.CreateCloudTableClient();
               this.GoodReadsCloudTable = cloudTableClient.GetTableReference(this.TableName);
               await this.GoodReadsCloudTable.CreateIfNotExistsAsync();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error occurred while creating the table.");
                throw;
            }
        }
    }
}
