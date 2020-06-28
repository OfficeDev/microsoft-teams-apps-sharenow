// <copyright file="UserPrivatePostController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Controller to handle user's private posts API operations.
    /// </summary>
    [Route("api/userprivatepost")]
    [ApiController]
    [Authorize]
    public class UserPrivatePostController : BaseGoodReadsController
    {
        /// <summary>
        /// Represents maximum number of private posts per user.
        /// </summary>
        private const int UserPrivatePostMaxCount = 50;

        /// <summary>
        /// Instance of Search service for working with storage.
        /// </summary>
        private readonly IPostSearchService postSearchService;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of user private post storage provider for private posts.
        /// </summary>
        private readonly IUserPrivatePostStorageProvider userPrivatePostStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserPrivatePostController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="userPrivatePostStorageProvider">User private post storage provider dependency injection.</param>
        /// <param name="postSearchService">The team post search service dependency injection.</param>
        public UserPrivatePostController(
            ILogger<UserPrivatePostController> logger,
            TelemetryClient telemetryClient,
            IUserPrivatePostStorageProvider userPrivatePostStorageProvider,
            IPostSearchService postSearchService)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.userPrivatePostStorageProvider = userPrivatePostStorageProvider;
            this.postSearchService = postSearchService;
        }

        /// <summary>
        /// Get call to retrieve list of private posts.
        /// </summary>
        /// <returns>List of private posts.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync()
        {
            try
            {
                this.logger.LogInformation("call to retrieve list of private posts.");

                var postIds = await this.userPrivatePostStorageProvider.GetUserPrivatePostsIdAsync(this.UserAadId);
                if (postIds != null || postIds.Any())
                {
                    var postIdsString = string.Join(";", postIds);
                    var filterQuery = $"search.in(PostId, '{postIdsString}', ';')";
                    var posts = await this.postSearchService.GetPostsAsync(PostSearchScope.SearchTeamPostsForTitleText, "*", userObjectId: null, filterQuery: filterQuery);
                    this.RecordEvent("Private posts - HTTP Get call succeeded");
                    return this.Ok(posts);
                }

                return this.Ok(new List<PostEntity>());
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to private post service.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store private posts details data.
        /// </summary>
        /// <param name="userPrivatePostEntity">Represents user private post entity object.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] UserPrivatePostEntity userPrivatePostEntity)
        {
            this.logger.LogInformation("Call to add private post.");

            try
            {
                var postIds = await this.userPrivatePostStorageProvider.GetUserPrivatePostsIdAsync(this.UserAadId);

                if (postIds.Count() < UserPrivatePostMaxCount)
                {
                    UserPrivatePostEntity userPrivatePost = new UserPrivatePostEntity
                    {
                        UserId = this.UserAadId,
                        CreatedByName = this.UserName,
#pragma warning disable CA1062 // private post details are validated by model validations for null check and is responded with bad request status
                        PostId = userPrivatePostEntity.PostId,
#pragma warning restore CA1062 // private post details are validated by model validations for null check and is responded with bad request status
                        CreatedDate = DateTime.UtcNow,
                    };

                    var result = await this.userPrivatePostStorageProvider.UpsertUserPrivatPostAsync(userPrivatePost);

                    if (result)
                    {
                        this.RecordEvent("Private posts - HTTP Post call succeeded");
                    }

                    return this.Ok(result);
                }
                else
                {
                    return this.Ok(false);
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while saving private post.");
                throw;
            }
        }

        /// <summary>
        /// Delete call to delete private post details data.
        /// </summary>
        /// <param name="postId">Id of the post to be deleted.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteAsync(string postId)
        {
            this.logger.LogInformation("call to delete private post.");

            if (string.IsNullOrEmpty(postId))
            {
                this.logger.LogError("PostId is either null or empty.");
                return this.BadRequest(new { message = "PostId is either null or empty." });
            }

            try
            {
                this.RecordEvent("Private posts - HTTP Delete call succeeded");
                return this.Ok(await this.userPrivatePostStorageProvider.DeleteUserPrivatePostAsync(postId, this.UserAadId));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while deleting private post {postId} for user {this.UserAadId}.");
                throw;
            }
        }
    }
}