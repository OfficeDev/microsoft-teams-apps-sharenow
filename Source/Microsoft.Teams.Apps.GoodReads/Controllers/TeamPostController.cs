// <copyright file="TeamPostController.cs" company="Microsoft">
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
    using Microsoft.Teams.Apps.GoodReads.Authentication;
    using Microsoft.Teams.Apps.GoodReads.Common;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Helpers;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Controller to handle team post API operations.
    /// </summary>
    [ApiController]
    [Route("api/teampost")]
    [Authorize]
    public class TeamPostController : BaseGoodReadsController
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
        /// Post search service for fetching post with search criteria and filters.
        /// </summary>
        private readonly IPostSearchService postSearchService;

        /// <summary>
        /// Provides method for fetching tags configured for team.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamPostController"/> class.
        /// </summary>
        /// <param name="logger">Used to perform logging of errors and information.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="postStorageHelper">Helper for creating models and filtering posts as per criteria.</param>
        /// <param name="postSearchService">Post search service for fetching post with search criteria and filters.</param>
        /// <param name="teamTagStorageProvider">Provides method for fetching tags configured for team.</param>
        public TeamPostController(
            ILogger<TeamPostController> logger,
            TelemetryClient telemetryClient,
            IPostStorageHelper postStorageHelper,
            IPostSearchService postSearchService,
            ITeamTagStorageProvider teamTagStorageProvider)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.postStorageHelper = postStorageHelper;
            this.postSearchService = postSearchService;
            this.teamTagStorageProvider = teamTagStorageProvider;
        }

        /// <summary>
        /// Get filtered team posts for particular team as per the configured tags.
        /// </summary>
        /// <param name="teamId">Team id for which data will fetch.</param>
        /// <param name="pageCount">Page number to get search data.</param>
        /// <returns>Returns filtered list of team posts as per the configured tags.</returns>
        [HttpGet("team-discover-posts")]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetTeamPostsAsync(string teamId, int pageCount)
        {
            this.logger.LogInformation("Call to get filtered team post details.");

            if (string.IsNullOrEmpty(teamId))
            {
                this.logger.LogError("TeamId is either null or empty.");
                return this.BadRequest(new { message = "TeamId is either null or empty." });
            }

            if (pageCount < 0)
            {
                this.logger.LogError("Invalid parameter value for pageCount.");
                return this.BadRequest(new { message = "Invalid parameter value for pageCount." });
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            try
            {
                // Get tags based on the team id for which tags has configured.
                var teamTagEntity = await this.teamTagStorageProvider.GetTeamTagAsync(teamId);

                if (teamTagEntity != null && !string.IsNullOrEmpty(teamTagEntity.Tags))
                {
                    // Prepare query based on the tags and get the data using search service.
                    var tagsQuery = this.postStorageHelper.GetTags(teamTagEntity.Tags);
                    var postsResult = await this.postSearchService.GetPostsAsync(PostSearchScope.FilterAsPerTeamTags, tagsQuery, userObjectId: null, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords);

                    if (postsResult != null && postsResult.Any())
                    {
                        // Filter the data based on the configured tags.
                        var filteredTeamPosts = this.postStorageHelper.GetFilteredTeamPostsAsPerTags(postsResult, teamTagEntity.Tags);
                        this.RecordEvent("Filtered team post - HTTP Get call succeeded");
                        return this.Ok(filteredTeamPosts);
                    }

                    this.logger.LogInformation($"No posts found for configured tags for team {teamId}.");
                }
                else
                {
                    this.logger.LogInformation($"Tags are not configured for team {teamId}.");
                }

                return this.Ok(new List<PostEntity>());
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while fetching posts for team {teamId}.");
                throw;
            }
        }

        /// <summary>
        /// Get posts for team according to filters.
        /// </summary>
        /// <param name="postTypes">Semicolon separated post type Ids. See more <see cref="PostTypeHelper"/>.</param>
        /// <param name="sharedByNames">Semicolon separated User names to filter the posts.</param>
        /// <param name="tags">Semicolon separated tags to match the post tags for which data will fetch.</param>
        /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="teamId">Team id to get configured tags for a team.</param>
        /// <param name="pageCount">Page count for which post needs to be fetched.</param>
        /// <returns>Returns filtered list of team posts as per the selected filters.</returns>
        [HttpGet("filtered-team-posts")]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetFilteredTeamPostsAsync(string postTypes, string sharedByNames, string tags, int sortBy, string teamId, int pageCount)
        {
            this.logger.LogInformation("Call to get team posts as per the applied filters.");

            if (pageCount < 0)
            {
                this.logger.LogError("Invalid argument value for pageCount.");
                return this.BadRequest(new { message = "Invalid argument value for pageCount." });
            }

            if (string.IsNullOrEmpty(teamId))
            {
                this.logger.LogError("Argument teamId cannot be null or empty.");
                return this.BadRequest(new { message = "Argument teamId cannot be null or empty." });
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;
            try
            {
                var teamTagEntity = await this.teamTagStorageProvider.GetTeamTagAsync(teamId);

                if (teamTagEntity == null || string.IsNullOrEmpty(teamTagEntity.Tags))
                {
                    this.logger.LogInformation($"Tags are not configured for team {teamId}.");
                    return this.BadRequest(new { message = $"Tags are not configured for team {teamId}." });
                }

                // If none of tags are selected for filtering, assign all configured tags for team to get posts which are intended for team.
                if (string.IsNullOrEmpty(tags))
                {
                    tags = teamTagEntity.Tags;
                }
                else
                {
                    var savedTags = teamTagEntity.Tags.Split(";");
                    var tagsList = tags.Split(';').Intersect(savedTags);
                    tags = string.Join(';', tagsList);
                }

                var tagsQuery = this.postStorageHelper.GetTags(tags);
                var filterQuery = this.postStorageHelper.GetFilterSearchQuery(postTypes, sharedByNames);
                var teamPosts = await this.postSearchService.GetPostsAsync(PostSearchScope.FilterTeamPosts, tagsQuery, userObjectId: null, sortBy: sortBy, filterQuery: filterQuery, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords);

                this.RecordEvent("Team post applied filters - HTTP Get call succeeded");
                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while getting posts using applied filters for team {teamId}.");
                throw;
            }
        }

        /// <summary>
        /// Get posts as per title search string.
        /// </summary>
        /// <param name="searchText">Search text represents the title of the posts.</param>
        /// <param name="teamId">Team Id for which posts needs to be fetched.</param>
        /// <param name="pageCount">Page count for which post needs to be fetched.</param>
        /// <returns>List of posts as per the title and configured tags.</returns>
        [HttpGet("search-posts")]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> SearchTeamPostsAsync(string searchText, string teamId, int pageCount)
        {
            this.logger.LogInformation("Call to get list of posts as per the configured tags and title.");

            if (string.IsNullOrEmpty(teamId))
            {
                this.logger.LogError("TeamId is either null or empty.");
                return this.BadRequest(new { message = "TeamId is either null or empty." });
            }

            if (pageCount < 0)
            {
                this.logger.LogError("Invalid argument value for pageCount.");
                return this.BadRequest(new { message = "Invalid argument value for pageCount." });
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;
            try
            {
                var teamTagEntity = await this.teamTagStorageProvider.GetTeamTagAsync(teamId);

                if (teamTagEntity == null || string.IsNullOrEmpty(teamTagEntity.Tags))
                {
                    this.logger.LogInformation($"Tags are not configured for team {teamId}.");
                    return this.BadRequest(new { message = $"Tags are not configured for team {teamId}." });
                }

                var tagsQuery = this.postStorageHelper.GetTags(teamTagEntity.Tags);
                var filterQuery = $"search.ismatch('{tagsQuery}', '{nameof(PostEntity.Tags)}')";
                var teamPosts = await this.postSearchService.GetPostsAsync(PostSearchScope.SearchTeamPostsForTitleText, searchText, userObjectId: null, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords, filterQuery: filterQuery);

                this.RecordEvent("Team post search result - HTTP Get call succeeded");
                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while getting posts as per search text for title for team {teamId}.");
                throw;
            }
        }

        /// <summary>
        /// Get unique author names.
        /// </summary>
        /// <param name="teamId">Team Id to get the configured tags for a team.</param>
        /// <returns>Returns unique user names.</returns>
        [HttpGet("team-post-authors")]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetAuthorNamesAsync(string teamId)
        {
            if (string.IsNullOrEmpty(teamId))
            {
                this.logger.LogError("TeamId is either null or empty.");
                return this.BadRequest(new { message = "TeamId is either null or empty." });
            }

            try
            {
                var authorNames = new List<string>();

                // Get tags based on the team id for which tags has configured.
                var teamTagEntity = await this.teamTagStorageProvider.GetTeamTagAsync(teamId);

                if (teamTagEntity == null || string.IsNullOrEmpty(teamTagEntity.Tags))
                {
                    this.logger.LogInformation($"Tags are not configured for team {teamId}.");
                    return this.Ok(authorNames);
                }

                var tagsQuery = this.postStorageHelper.GetTags(teamTagEntity.Tags);
                var posts = await this.postSearchService.GetPostsAsync(PostSearchScope.FilterAsPerTeamTags, tagsQuery, userObjectId: null);

                if (posts != null)
                {
                    authorNames = this.postStorageHelper.GetAuthorNamesAsync(posts).ToList();
                    this.RecordEvent("Team post unique author names - HTTP Get call succeeded");
                }

                return this.Ok(authorNames);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique user names.");
                throw;
            }
        }
    }
}