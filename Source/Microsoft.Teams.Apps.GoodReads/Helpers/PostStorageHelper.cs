// <copyright file="PostStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Implements post storage helper which helps to construct the model, create search query for post.
    /// </summary>
    public class PostStorageHelper : IPostStorageHelper
    {
        /// <summary>
        /// Get filtered posts as per the configured tags.
        /// </summary>
        /// <param name="teamPosts">Post entities.</param>
        /// <param name="teamConfiguredTags">Search text for tags.</param>
        /// <returns>Represents team posts.</returns>
        public IEnumerable<PostEntity> GetFilteredTeamPostsAsPerTags(IEnumerable<PostEntity> teamPosts, string teamConfiguredTags)
        {
            var filteredTeamPosts = new List<PostEntity>();
#pragma warning disable CA1062 // Validating argument in caller
            var searchTagList = teamConfiguredTags.ToUpperInvariant().Split(";");
            foreach (var teamPost in teamPosts)
#pragma warning restore CA1062 // Validating argument in caller
            {
                if (!string.IsNullOrEmpty(teamPost.Tags))
                {
                    var postTags = teamPost.Tags.ToUpperInvariant().Split(";");
                    if (searchTagList.Intersect(postTags).Any())
                    {
                        filteredTeamPosts.Add(teamPost);
                    }
                }
            }

            return filteredTeamPosts;
        }

        /// <summary>
        /// Get tags to fetch posts as per the configured tags.
        /// </summary>
        /// <param name="tags">Tags of a configured post.</param>
        /// <returns>Represents tags to fetch posts.</returns>
        public string GetTags(string tags)
        {
#pragma warning disable CA1062 // Validating argument in caller
            var postTags = tags.Split(';').Where(postType => !string.IsNullOrWhiteSpace(postType));
#pragma warning restore CA1062 // Validating argument in caller
            return string.Join(" ", postTags);
        }

        /// <summary>
        /// Get filtered posts as per the date range.
        /// </summary>
        /// <param name="posts">Posts data.</param>
        /// <param name="fromDate">Start date from which data should fetch.</param>
        /// <param name="toDate">End date till when data should fetch.</param>
        /// <returns>A task that represent collection to hold posts data.</returns>
        public IEnumerable<PostEntity> GetTeamPostsForDateRange(IEnumerable<PostEntity> posts, DateTime fromDate, DateTime toDate)
        {
            return posts.Where(post => post.UpdatedDate >= fromDate && post.UpdatedDate <= toDate);
        }

        /// <summary>
        /// Get filtered 50 user names from posts data.
        /// </summary>
        /// <param name="posts">Represents a collection of posts.</param>
        /// <returns>Represents posts.</returns>
        public IEnumerable<string> GetAuthorNamesAsync(IEnumerable<PostEntity> posts)
        {
            return posts
                .GroupBy(post => post.UserId)
                .OrderByDescending(groupedPost => groupedPost.Count())
                .Take(50)
                .Select(post => post.First().CreatedByName)
                .OrderBy(createdByName => createdByName);
        }

        /// <summary>
        /// Get combined query to fetch posts as per the selected filter.
        /// </summary>
        /// <param name="postTypes">Post type see <see cref="PostTypeHelper"/>.</param>
        /// <param name="sharedByNames">User names selected in filter.</param>
        /// <returns>Represents user names query to filter posts.</returns>
        public string GetFilterSearchQuery(string postTypes, string sharedByNames)
        {
#pragma warning disable CA1062 // Validating argument in caller
            var typesQuery = string.IsNullOrEmpty(postTypes) ? null : this.GetPostTypesQuery(postTypes);
            var sharedByNamesQuery = string.IsNullOrEmpty(sharedByNames) ? null : this.GetSharedByNamesQuery(sharedByNames);
#pragma warning restore CA1062 // Validating argument in caller
            if (string.IsNullOrEmpty(typesQuery) && string.IsNullOrEmpty(sharedByNamesQuery))
            {
                return null;
            }

            if (!string.IsNullOrEmpty(typesQuery) && !string.IsNullOrEmpty(sharedByNamesQuery))
            {
                return $"({typesQuery}) and ({sharedByNamesQuery})";
            }

            if (!string.IsNullOrEmpty(typesQuery))
            {
                return $"({typesQuery})";
            }

            if (!string.IsNullOrEmpty(sharedByNamesQuery))
            {
                return $"({sharedByNamesQuery})";
            }

            return null;
        }

        /// <summary>
        /// Get post type query to fetch posts as per the selected filter.
        /// </summary>
        /// <param name="postTypes">Post type see <see cref="PostTypeHelper"/>.</param>
        /// <returns>Represents post type query to filter posts.</returns>
        private string GetPostTypesQuery(string postTypes)
        {
            StringBuilder postTypesQuery = new StringBuilder();
            var postTypesData = postTypes.Split(';').Where(postType => !string.IsNullOrWhiteSpace(postType)).Select(postType => postType.Trim());

            if (postTypesData.Count() > 1)
            {
                var posts = postTypesData.Take(postTypesData.Count() - 1).ToList();
                posts.ForEach(postType =>
                {
                    postTypesQuery.Append($"Type eq {postType} or ");
                });

                postTypesQuery.Append($"Type eq {postTypesData.Last()}");
            }
            else
            {
                postTypesQuery.Append($"Type eq {postTypesData.Last()}");
            }

            return postTypesQuery.ToString();
        }

        /// <summary>
        /// Get user names query to fetch posts as per the selected filter.
        /// </summary>
        /// <param name="sharedByNames">User names selected in filter.</param>
        /// <returns>Represents user names query to filter posts.</returns>
        private string GetSharedByNamesQuery(string sharedByNames)
        {
            StringBuilder sharedByNamesQuery = new StringBuilder();
            var sharedByNamesData = sharedByNames.Split(';').Where(name => !string.IsNullOrWhiteSpace(name)).Select(name => name.Trim());

            if (sharedByNamesData.Count() > 1)
            {
                var users = sharedByNamesData.Take(sharedByNamesData.Count() - 1).ToList();
                users.ForEach(user =>
                {
                    sharedByNamesQuery.Append($"CreatedByName eq '{user}' or ");
                });

                sharedByNamesQuery.Append($"CreatedByName eq '{sharedByNamesData.Last()}'");
            }
            else
            {
                sharedByNamesQuery.Append($"CreatedByName eq '{sharedByNamesData.Last()}'");
            }

            return sharedByNamesQuery.ToString();
        }
    }
}
