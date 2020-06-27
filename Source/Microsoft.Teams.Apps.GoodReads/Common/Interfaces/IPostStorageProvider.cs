// <copyright file="IPostStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.GoodReads.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Interface for provider which helps in retrieving, storing, updating and deleting post details.
    /// </summary>
    public interface IPostStorageProvider
    {
        /// <summary>
        /// Get posts data.
        /// </summary>
        /// <param name="isRemoved">Represent whether a post is deleted or not.</param>
        /// <returns>A task that represent collection to hold posts ids.</returns>
        Task<IEnumerable<PostEntity>> GetPostsAsync(bool isRemoved);

        /// <summary>
        /// Stores or update post details data.
        /// </summary>
        /// <param name="postEntity">Holds post detail entity data.</param>
        /// <returns>A task that represents post entity data is saved or updated.</returns>
        Task<bool> UpsertPostAsync(PostEntity postEntity);

        /// <summary>
        /// Get post data.
        /// </summary>
        /// <param name="postCreatedByuserId">User id to fetch the post details.</param>
        /// <param name="postId">Post id to fetch the post details.</param>
        /// <returns>A task that represent a object to hold post data.</returns>
        Task<PostEntity> GetPostAsync(string postCreatedByuserId, string postId);

        /// <summary>
        /// Get posts as per the user's private list.
        /// </summary>
        /// <param name="postIds">A collection of user private post id's.</param>
        /// <returns>A task that represent collection to hold posts data.</returns>
        Task<IEnumerable<PostEntity>> GetFilteredUserPrivatePostsAsync(IEnumerable<string> postIds);
    }
}