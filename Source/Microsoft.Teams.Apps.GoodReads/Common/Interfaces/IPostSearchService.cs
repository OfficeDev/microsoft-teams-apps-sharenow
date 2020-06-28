// <copyright file="IPostSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Interface for post search service which helps in searching posts using Azure Search service.
    /// </summary>
    public interface IPostSearchService
    {
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
        Task<IEnumerable<PostEntity>> GetPostsAsync(
            PostSearchScope searchScope,
            string searchQuery,
            string userObjectId,
            int? count = null,
            int? skip = null,
            int? sortBy = null,
            string filterQuery = null);

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task RecreateSearchServiceIndexAsync();

        /// <summary>
        /// Run the indexer on demand.
        /// </summary>
        /// <returns>A task that represents the work queued to execute</returns>
        Task RunIndexerOnDemandAsync();
    }
}
