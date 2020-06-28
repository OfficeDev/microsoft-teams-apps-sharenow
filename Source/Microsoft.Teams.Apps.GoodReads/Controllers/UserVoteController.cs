// <copyright file="UserVoteController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Microsoft.WindowsAzure.Storage;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Controller to handle user vote operations.
    /// </summary>
    [ApiController]
    [Route("api/uservote")]
    [Authorize]
    public class UserVoteController : BaseGoodReadsController
    {
        /// <summary>
        /// Retry policy with jitter.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private readonly AsyncRetryPolicy retryPolicy;

        /// <summary>
        /// Used to perform logging of errors and information.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Provider for working with user vote data in storage.
        /// </summary>
        private readonly IUserVoteStorageProvider userVoteStorageProvider;

        /// <summary>
        /// Provider to fetch posts from storage.
        /// </summary>
        private readonly IPostStorageProvider postStorageProvider;

        /// <summary>
        /// Search service instance for fetching posts using filters and search queries.
        /// </summary>
        private readonly IPostSearchService postSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserVoteController"/> class.
        /// </summary>
        /// <param name="logger">Used to perform logging of errors and information.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="userVoteStorageProvider">Provider for working with user vote data in storage.</param>
        /// <param name="postStorageProvider">Provider to fetch posts from storage.</param>
        /// <param name="postSearchService">Search service instance for fetching posts using filters and search queries.</param>
        public UserVoteController(
            ILogger<TeamPostController> logger,
            TelemetryClient telemetryClient,
            IUserVoteStorageProvider userVoteStorageProvider,
            IPostStorageProvider postStorageProvider,
            IPostSearchService postSearchService)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.userVoteStorageProvider = userVoteStorageProvider;
            this.postStorageProvider = postStorageProvider;
            this.postSearchService = postSearchService;
            this.retryPolicy = Policy.Handle<StorageException>(ex => ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
                .WaitAndRetryAsync(Backoff.LinearBackoff(TimeSpan.FromMilliseconds(250), 25));
        }

        /// <summary>
        /// Retrieves list of votes for user.
        /// </summary>
        /// <returns>List of posts.</returns>
        [HttpGet("user-votes")]
        public async Task<IActionResult> GetVotesAsync()
        {
            try
            {
                this.logger.LogInformation("call to retrieve list of votes for user.");

                var userVotes = await this.userVoteStorageProvider.GetUserVotesAsync(this.UserAadId);
                this.RecordEvent("User votes - HTTP Get call succeeded.");

                return this.Ok(userVotes);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to post service.");
                throw;
            }
        }

        /// <summary>
        /// Stores user vote for a post.
        /// </summary>
        /// <param name="postCreatedByUserId">AAD user Id of user who created post.</param>
        /// <param name="postId">Id of the post to delete vote.</param>
        /// <remarks> Note: the implementation here uses Azure table storage for handling votes
        /// in posts and user vote tables. Table storage are not transactional and there
        /// can be instances where the vote count might be off. The table operations are
        /// wrapped with retry policies in case of conflict or failures to minimize the risks.</remarks>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("vote")]
        public async Task<IActionResult> AddVoteAsync(string postCreatedByUserId, string postId)
        {
            this.logger.LogInformation("call to add user vote.");

            if (string.IsNullOrEmpty(postCreatedByUserId))
            {
                this.logger.LogError("Error while deleting vote. Parameter postCreatedByuserId is either null or empty.");
                return this.BadRequest(new { message = "Parameter postCreatedByuserId is either null or empty." });
            }

            if (string.IsNullOrEmpty(postId))
            {
                this.logger.LogError("Error while deleting vote. PostId is either null or empty.");
                return this.BadRequest(new { message = "PostId is either null or empty." });
            }

            bool isUserVoteSavedSuccessful = false;
            bool isPostSavedSuccessful = false;

            try
            {
#pragma warning disable CA1062 // post details are validated by model validations for null check and is responded with bad request status
                var userVoteForPost = await this.userVoteStorageProvider.GetUserVoteForPostAsync(this.UserAadId, postId);
#pragma warning restore CA1062 // post details are validated by model validations for null check and is responded with bad request status

                if (userVoteForPost == null)
                {
                    UserVoteEntity userVote = new UserVoteEntity
                    {
                        UserId = this.UserAadId,
                        PostId = postId,
                    };

                    await this.retryPolicy.ExecuteAsync(async () =>
                    {
                        isUserVoteSavedSuccessful = await this.AddUserVoteAsync(userVote);
                    });

                    if (!isUserVoteSavedSuccessful)
                    {
                        this.logger.LogError($"User vote is not updated successfully for post {postId} by {this.UserAadId} ");
                        return this.StatusCode(StatusCodes.Status500InternalServerError, "An error occurred while saving user vote.");
                    }

                    // Retry if storage operation conflict occurs during updating user vote count.
                    await this.retryPolicy.ExecuteAsync(async () =>
                    {
                        isPostSavedSuccessful = await this.UpdateTotalCountAsync(postCreatedByUserId, postId, isUpvote: true);
                    });
                }
            }
#pragma warning disable CA1031 // catching generic exception to trace error in telemetry and return false value to client
            catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace error in telemetry and return false value to client
            {
                this.logger.LogError(ex, "Exception occurred while updating user vote.");
            }
            finally
            {
                if (isPostSavedSuccessful)
                {
                    // run Azure search service to refresh the index for getting latest vote count
                    await this.postSearchService.RunIndexerOnDemandAsync();
                }
                else
                {
                    // revert user vote entry if the post total count didn't saved successfully
                    this.logger.LogError($"Post vote count is not updated successfully for post {postId} by {this.UserAadId} ");

                    // exception handling is implemented in method and no additional check is required
                    var isUserVoteDeletedSuccessful = await this.userVoteStorageProvider.DeleteUserVoteAsync(postId, this.UserAadId);
                    if (isUserVoteDeletedSuccessful)
                    {
                        this.logger.LogInformation("Vote revoked from user table");
                    }
                    else
                    {
                        this.logger.LogError("Vote cannot be revoked from user table");
                    }
                }
            }

            return this.Ok(isPostSavedSuccessful);
        }

        /// <summary>
        /// Deletes user vote for a post.
        /// </summary>
        /// <param name="postCreatedByUserId">AAD user Id of user who created post.</param>
        /// <param name="postId">Id of the post to delete vote.</param>
        /// <remarks> Note: the implementation here uses Azure table storage for handling votes
        /// in posts and user vote tables. Table storage are not transactional and there
        /// can be instances where the vote count might be off. The table operations are
        /// wrapped with retry policies in case of conflict or failures to minimize the risks.</remarks>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteVoteAsync(string postCreatedByUserId, string postId)
        {
            this.logger.LogInformation("call to delete user vote.");

            if (string.IsNullOrEmpty(postCreatedByUserId))
            {
                this.logger.LogError("Error while deleting vote. Parameter postCreatedByuserId is either null or empty.");
                return this.BadRequest(new { message = "Parameter postCreatedByuserId is either null or empty." });
            }

            if (string.IsNullOrEmpty(postId))
            {
                this.logger.LogError("Error while deleting vote. PostId is either null or empty.");
                return this.BadRequest(new { message = "PostId is either null or empty." });
            }

            bool isPostSavedSuccessful = false;
            bool isUserVoteDeletedSuccessful = false;

            try
            {
                isUserVoteDeletedSuccessful = await this.userVoteStorageProvider.DeleteUserVoteAsync(postId, this.UserAadId);

                if (!isUserVoteDeletedSuccessful)
                {
                    this.logger.LogError($"Vote is not updated successfully for post {postId} by {postCreatedByUserId} ");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Vote is not updated successfully.");
                }

                // Retry if storage operation conflict occurs while updating post count.
                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    isPostSavedSuccessful = await this.UpdateTotalCountAsync(postCreatedByUserId, postId, isUpvote: false);
                });
            }
#pragma warning disable CA1031 // catching generic exception to trace error in telemetry and return false value to client
            catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace error in telemetry and return false value to client
            {
                this.logger.LogError(ex, "Exception occured while deleting the user vote count.");
            }
            finally
            {
                // if user vote is not saved successfully
                // revert back the total post count
                if (isPostSavedSuccessful)
                {
                    // run Azure search service to refresh the index for getting latest vote count
                    await this.postSearchService.RunIndexerOnDemandAsync();
                }
                else
                {
                    UserVoteEntity userVote = new UserVoteEntity
                    {
                        UserId = this.UserAadId,
                        PostId = postId,
                    };

                    // add the user vote back to table
                    await this.retryPolicy.ExecuteAsync(async () =>
                    {
                         await this.AddUserVoteAsync(userVote);
                    });
                }
            }

            return this.Ok(isPostSavedSuccessful);
        }

        /// <summary>
        /// Add user vote in store with retry attempts
        /// </summary>
        /// <param name="userVote">User vote instance with user and post id</param>
        /// <returns>True if operation executed successfully else false</returns>
        private async Task<bool> AddUserVoteAsync(UserVoteEntity userVote)
        {
            bool isUserVoteSavedSuccessful = false;
            try
            {
                // Update operation will throw exception if the column has already been updated
                // or if there is a transient error (handled by an Azure storage internally)
                isUserVoteSavedSuccessful = await this.userVoteStorageProvider.UpsertUserVoteAsync(userVote);
            }
            catch (StorageException ex)
            {
                if (ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
                {
                    this.logger.LogInformation("Optimistic concurrency violation – entity has changed since it was retrieved.");
                    throw;
                }
            }
#pragma warning disable CA1031 // catching generic exception to trace log error in telemetry and continue the execution
            catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace log error in telemetry and continue the execution
            {
                // log exception details to telemetry
                // but do not attempt to retry in order to avoid multiple vote count decrement
                this.logger.LogError(ex, "Exception occurred while reading post details.");
            }

            return isUserVoteSavedSuccessful;
        }

        /// <summary>
        /// Increement or decreement the total vote counts of post
        /// </summary>
        /// <param name="postCreatedByUserId">Post owner user object id</param>
        /// <param name="postId">Post unique id</param>
        /// <param name="isUpvote">Set true to increase total count by 1 else false</param>
        /// <returns>True if operation exectuted successfully else false</returns>
        private async Task<bool> UpdateTotalCountAsync(string postCreatedByUserId, string postId, bool isUpvote = false)
        {
            bool isPostSavedSuccessful = false;
            try
            {
                var postEntity = await this.postStorageProvider.GetPostAsync(postCreatedByUserId, postId);

                postEntity.TotalVotes = isUpvote ? postEntity.TotalVotes + 1 : postEntity.TotalVotes - 1;

                if (postEntity.TotalVotes >= 0)
                {
                    isPostSavedSuccessful = await this.postStorageProvider.UpsertPostAsync(postEntity);
                }
            }
            catch (StorageException ex)
            {
                if (ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
                {
                    this.logger.LogInformation("Optimistic concurrency violation – entity has changed since it was retrieved.");
                    throw;
                }
            }
#pragma warning disable CA1031 // catching generic exception to trace log error in telemetry and continue the execution
            catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace log error in telemetry and continue the execution
            {
                // log exception details to telemetry
                // but do not attempt to retry in order to avoid multiple vote count increment
                this.logger.LogError(ex, "Exception occurred while reading post details.");
            }

            return isPostSavedSuccessful;
        }
    }
}