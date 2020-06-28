// <copyright file="PostController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GoodReads.Common;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Helpers;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Controller to handle post API operations.
    /// </summary>
    [ApiController]
    [Route("api/userposts")]
    [Authorize]
    public class PostController : BaseGoodReadsController
    {
        /// <summary>
        /// Used to perform logging of errors and information.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Helper for creating models and filtering posts as per criteria.
        /// </summary>
        private readonly IPostStorageHelper postStorageHelper;

        /// <summary>
        /// Provides methods for add, update and delete post operations from database.
        /// </summary>
        private readonly IPostStorageProvider postStorageProvider;

        /// <summary>
        /// Post search service for fetching post with search criteria and filters.
        /// </summary>
        private readonly IPostSearchService postSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="PostController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="postStorageHelper">Helper for creating models and filtering posts as per criteria.</param>
        /// <param name="postStorageProvider">Provides methods for add, update and delete post operations from database.</param>
        /// <param name="postSearchService">Post search service for fetching post with search criteria and filters.</param>
        public PostController(
            ILogger<TeamPostController> logger,
            TelemetryClient telemetryClient,
            IPostStorageHelper postStorageHelper,
            IPostStorageProvider postStorageProvider,
            IPostSearchService postSearchService)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.postStorageHelper = postStorageHelper;
            this.postStorageProvider = postStorageProvider;
            this.postSearchService = postSearchService;
        }

        /// <summary>
        /// Fetch posts according to page count.
        /// </summary>
        /// <param name="pageCount">Page number to get search data.</param>
        /// <returns>List of posts.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync(int pageCount)
        {
            this.logger.LogInformation("Call to retrieve list of posts.");

            if (pageCount < 0)
            {
                this.logger.LogError("Invalid value for argument pageCount.");
                return this.BadRequest(new { message = "Invalid value for argument pageCount." });
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            try
            {
                var posts = await this.postSearchService.GetPostsAsync(PostSearchScope.AllItems, searchQuery: null, userObjectId: null, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords);
                this.RecordEvent("Get post - HTTP Get call succeeded");

                return this.Ok(posts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching posts.");
                throw;
            }
        }

        /// <summary>
        /// Stores new post details.
        /// </summary>
        /// <param name="postDetails">Post detail which needs to be stored.</param>
        /// <returns>Returns added post for successful operation or false for failure.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] PostEntity postDetails)
        {
            this.RecordEvent("Save post - HTTP Post call initiated");

            try
            {
                var updatedPostEntity = new PostEntity
                {
#pragma warning disable CA1062 // post details are validated by model validations for null check and is responded with bad request status
                    ContentUrl = postDetails.ContentUrl,
#pragma warning restore CA1062 // post detail are validated by model validations for null check and is responded with bad request status
                    CreatedByName = this.UserName,
                    CreatedDate = DateTime.UtcNow,
                    Description = postDetails.Description,
                    IsRemoved = false,
                    PostId = Guid.NewGuid().ToString(),
                    Tags = postDetails.Tags,
                    Title = postDetails.Title,
                    TotalVotes = 0,
                    Type = postDetails.Type,
                    UpdatedDate = DateTime.UtcNow,
                    UserId = this.UserAadId,
                };

                var result = await this.postStorageProvider.UpsertPostAsync(updatedPostEntity);

                // If operation is successful, run Azure search service indexer.
                if (result)
                {
                    this.RecordEvent("Save post - HTTP Post call succeeded");
                    await this.postSearchService.RunIndexerOnDemandAsync();
                    return this.Ok(updatedPostEntity);
                }
                else
                {
                    this.RecordEvent("Save post - HTTP Post call failed");
                    return this.Ok(false);
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while adding new post.");
                throw;
            }
        }

        /// <summary>
        /// Get team posts as per the applied filters.
        /// </summary>
        /// <param name="postTypes">Semicolon separated types of posts. See more <see cref="PostTypeHelper"/>.</param>
        /// /// <param name="sharedByNames">Semicolon separated User names to filter the posts.</param>
        /// /// <param name="tags">Semicolon separated tags to match the post tags for which data will fetch.</param>
        /// /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="pageCount">Page count for which post needs to be fetched.</param>
        /// <returns>Returns filtered list of team posts as per the selected filters.</returns>
        [HttpGet("filtered-posts")]
        public async Task<IActionResult> GetFilteredPostsAsync(string postTypes, string sharedByNames, string tags, int sortBy, int pageCount)
        {
            this.logger.LogInformation("Call to get posts as per the applied filters.");

            if (pageCount < 0)
            {
                this.logger.LogError("Invalid argument value for pageCount.");
                return this.BadRequest(new { message = "Invalid value for argument pageCount." });
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;
            try
            {
                // If no tags selected for filtering then get posts irrespective of tags.
                var tagsQuery = string.IsNullOrEmpty(tags) ? "*" : this.postStorageHelper.GetTags(tags);
                var filterQuery = this.postStorageHelper.GetFilterSearchQuery(postTypes, sharedByNames);
                var teamPosts = await this.postSearchService.GetPostsAsync(PostSearchScope.FilterTeamPosts, tagsQuery, userObjectId: null, sortBy: sortBy, filterQuery: filterQuery, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords);

                this.RecordEvent("Team post applied filters - HTTP Get call succeeded");

                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching filtered posts.");
                throw;
            }
        }

        /// <summary>
        /// Updates existing post details.
        /// </summary>
        /// <param name="postDetails">Post details which needs to be updated.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPatch]
        public async Task<IActionResult> PatchAsync([FromBody] PostEntity postDetails)
        {
            this.logger.LogInformation("Call to update post details.");

#pragma warning disable CA1062 // Validation is done using data annotations
            if (string.IsNullOrEmpty(postDetails.PostId))
#pragma warning restore CA1062 // Validation is done using data annotations
            {
                this.logger.LogError($"PostId is either null or empty.");
                this.RecordEvent("Update post - HTTP Put call failed");

                return this.BadRequest(new { message = "PostId cannot be null or empty." });
            }

            if (postDetails.UserId != this.UserAadId)
            {
                this.logger.LogError($"User {this.UserAadId} did not create any post with post Id: {postDetails.PostId}.");
                this.RecordEvent("Update post - HTTP Put call failed");

                return this.NotFound(new { message = $"User did not create any post with post Id: {postDetails.PostId}." });
            }

            try
            {
                // Validating Post Id as it will be generated at server side in case of adding new post but cannot be null or empty in case of update.
                var currentPost = await this.postStorageProvider.GetPostAsync(this.UserAadId, postDetails.PostId);

                if (currentPost == null || currentPost.IsRemoved)
                {
                    this.logger.LogInformation($"Could not find post {postDetails.PostId} created by user {this.UserAadId}");
                    this.RecordEvent("Update post - HTTP Put call failed");

                    return this.BadRequest(new { message = $"Could not find post {postDetails.PostId} to update." });
                }

                currentPost.Description = postDetails.Description;
                currentPost.Tags = postDetails.Tags;
                currentPost.Title = postDetails.Title;
                currentPost.Type = postDetails.Type;
                currentPost.ContentUrl = postDetails.ContentUrl;

                var upsertResult = await this.postStorageProvider.UpsertPostAsync(currentPost);

                // If operation is successful, run indexer.
                if (upsertResult)
                {
                    this.RecordEvent("Update post - HTTP Put call succeeded");
                    await this.postSearchService.RunIndexerOnDemandAsync();
                }
                else
                {
                    this.RecordEvent("Update post - HTTP Put call failed");
                    this.logger.LogError("Update post action failed");
                }

                return this.Ok(upsertResult);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while updating post.");
                throw;
            }
        }

        /// <summary>
        /// Delete call to delete post details.
        /// </summary>
        /// <param name="postId">Post Id of the post to be deleted.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteAsync(string postId)
        {
            this.logger.LogInformation("Call to delete post details.");

            if (string.IsNullOrEmpty(postId))
            {
                this.logger.LogError("PostId is found null or empty while deleting the post.");
                return this.BadRequest(new { message = "PostId cannot be null or empty." });
            }

            try
            {
                var postDetails = await this.postStorageProvider.GetPostAsync(this.UserAadId, postId);

                if (postDetails == null)
                {
                    this.logger.LogError($"Post {postId} not found for deletion.");
                    return this.BadRequest(new { message = $"Cannot find post {postId} created by user {this.UserAadId} for deletion." });
                }

                postDetails.IsRemoved = true;
                var deletionResult = await this.postStorageProvider.UpsertPostAsync(postDetails);

                // Run indexer is operation is successful.
                if (deletionResult)
                {
                    await this.postSearchService.RunIndexerOnDemandAsync();
                    this.RecordEvent("Delete post - HTTP Delete call succeeded");
                }
                else
                {
                    this.RecordEvent("Delete post - HTTP Delete call failed");
                }

                return this.Ok(deletionResult);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while deleting post.");
                throw;
            }
        }

        /// <summary>
        /// Get unique user names.
        /// </summary>
        /// <returns>Returns unique user names.</returns>
        [HttpGet("unique-user-names")]
        public async Task<IActionResult> GetUniqueUserNamesAsync()
        {
            try
            {
                this.logger.LogInformation("Call to get unique names.");

                // Search query will be null if there is no search criteria used. userObjectId will be used when we want to get posts created by respective user.
                var postDetails = await this.postSearchService.GetPostsAsync(PostSearchScope.UniqueUserNames, searchQuery: null, userObjectId: null);

                if (postDetails == null)
                {
                    this.logger.LogInformation("No posts are available for search");

                    // return empty list with 200 status if no posts are added in storage yet.
                    this.Ok(new List<string>());
                }

                var authorNames = this.postStorageHelper.GetAuthorNamesAsync(postDetails);

                this.RecordEvent("User post unique user names - HTTP Get call succeeded");

                return this.Ok(authorNames);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching unique user names.");
                throw;
            }
        }

        /// <summary>
        /// Get list of posts as per the title text.
        /// </summary>
        /// <param name="searchText">Search text represents the title field to find and get posts.</param>
        /// <param name="pageCount">Page number to get search data from Azure Search service.</param>
        /// <returns>List of filtered posts as per the search text for title.</returns>
        [HttpGet("search-posts")]
        public async Task<IActionResult> GetSearchedPostsForTitleAsync(string searchText, int pageCount)
        {
            this.logger.LogInformation("Call to get list of posts according to searched post title text.");

            if (pageCount < 0)
            {
                this.logger.LogError("PageCount found to be less than or equal to zero while searching posts");
                return this.BadRequest(new { message = "PageCount cannot be less than or equal to zero." });
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            try
            {
                var searchedPosts = await this.postSearchService.GetPostsAsync(PostSearchScope.SearchTeamPostsForTitleText, searchText, userObjectId: null, skip: skipRecords, count: Constants.LazyLoadPerPagePostCount);
                this.RecordEvent("User post title search - HTTP Get call succeeded");

                return this.Ok(searchedPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching posts according to searched post title text.");
                throw;
            }
        }
    }
}