// <copyright file="PostSearchScope.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Models
{
    /// <summary>
    /// A enum that represent the search scope for search service.
    /// </summary>
    public enum PostSearchScope
    {
        /// <summary>
        /// Represents all team posts.
        /// </summary>
        AllItems,

        /// <summary>
        /// Represents posts created by current user.
        /// </summary>
        PostedByMe,

        /// <summary>
        /// Represents popular posts based on the number of votes.
        /// </summary>
        Popular,

        /// <summary>
        /// Represents configured team tags.
        /// </summary>
        TeamPreferenceTags,

        /// <summary>
        /// Represents filtered posts as per the configured tags in a particular team.
        /// </summary>
        FilterAsPerTeamTags,

        /// <summary>
        /// Represents posts based on the date range to send digest notification.
        /// </summary>
        FilterPostsAsPerDateRange,

        /// <summary>
        /// Represents unique user names who created the posts to show on filter bar drop-down list.
        /// </summary>
        UniqueUserNames,

        /// <summary>
        /// Represents posts as per the search text for title field.
        /// </summary>
        SearchTeamPostsForTitleText,

        /// <summary>
        /// Represents posts as per the applied filters.
        /// </summary>
        FilterTeamPosts,
    }
}
