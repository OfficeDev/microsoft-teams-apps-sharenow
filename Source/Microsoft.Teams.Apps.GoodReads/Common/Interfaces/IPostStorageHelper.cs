// <copyright file="IPostStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Interfaces
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.GoodReads.Helpers;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Interface for storage helper which helps in preparing model data for post.
    /// </summary>
    public interface IPostStorageHelper
    {
        /// <summary>
        /// Get filtered posts as per the configured tags.
        /// </summary>
        /// <param name="teamPosts">Team post entities.</param>
        /// <param name="teamConfiguredTags">Search text for tags.</param>
        /// <returns>Represents team posts.</returns>
        IEnumerable<PostEntity> GetFilteredTeamPostsAsPerTags(IEnumerable<PostEntity> teamPosts, string teamConfiguredTags);

        /// <summary>
        /// Get tags query to fetch posts as per the configured tags.
        /// </summary>
        /// <param name="tags">Tags of a configured post.</param>
        /// <returns>Represents tags query to fetch posts.</returns>
        string GetTags(string tags);

        /// <summary>
        /// Get filtered posts as per the date range.
        /// </summary>
        /// <param name="posts">Posts data.</param>
        /// <param name="fromDate">Start date from which data should fetch.</param>
        /// <param name="toDate">End date till when data should fetch.</param>
        /// <returns>A task that represent collection to hold posts data.</returns>
        IEnumerable<PostEntity> GetTeamPostsForDateRange(IEnumerable<PostEntity> posts, DateTime fromDate, DateTime toDate);

        /// <summary>
        /// Get filtered unique user names.
        /// </summary>
        /// <param name="posts">Posts details.</param>
        /// <returns>List of unique author names.</returns>
        IEnumerable<string> GetAuthorNamesAsync(IEnumerable<PostEntity> posts);

        /// <summary>
        /// Get combined query to fetch posts as per the selected filter.
        /// </summary>
        /// <param name="postTypes">Post type. see <see cref="PostTypeHelper"/>.</param>
        /// <param name="sharedByNames">User names selected in filter.</param>
        /// <returns>Represents user names query to filter posts.</returns>
        string GetFilterSearchQuery(string postTypes, string sharedByNames);
    }
}
