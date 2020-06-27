// <copyright file="TeamPreferenceController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GoodReads.Authentication;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Controller to handle team preference API operations.
    /// </summary>
    [Route("api/teampreference")]
    [ApiController]
    [Authorize]
    public class TeamPreferenceController : BaseGoodReadsController
    {
        /// <summary>
        /// Used to perform logging of errors and information.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Helper to create model for adding team preference and filtering unique tags.
        /// </summary>
        private readonly ITeamPreferenceStorageHelper teamPreferenceStorageHelper;

        /// <summary>
        /// Provider having methods to add and get team preferences from database.
        /// </summary>
        private readonly ITeamPreferenceStorageProvider teamPreferenceStorageProvider;

        /// <summary>
        /// Search service for fetching posts as per criteria.
        /// </summary>
        private readonly IPostSearchService teamPostSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamPreferenceController"/> class.
        /// </summary>
        /// <param name="logger">Used to perform logging of errors and information.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="teamPreferenceStorageHelper">Helper to create model for adding team preference and filtering unique tags.</param>
        /// <param name="teamPreferenceStorageProvider">Provider having methods to add and get team preferences from database.</param>
        /// <param name="teamPostSearchService">Search service for fetching posts as per criteria.</param>
        public TeamPreferenceController(
            ILogger<TeamPreferenceController> logger,
            TelemetryClient telemetryClient,
            ITeamPreferenceStorageHelper teamPreferenceStorageHelper,
            ITeamPreferenceStorageProvider teamPreferenceStorageProvider,
            IPostSearchService teamPostSearchService)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.teamPreferenceStorageHelper = teamPreferenceStorageHelper;
            this.teamPreferenceStorageProvider = teamPreferenceStorageProvider;
            this.teamPostSearchService = teamPostSearchService;
        }

        /// <summary>
        /// Get call to retrieve team preference data.
        /// </summary>
        /// <param name="teamId">Team id - unique value for each Team where preference has configured.</param>
        /// <returns>Returns team preference details.</returns>
        [HttpGet]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetAsync(string teamId)
        {
            this.logger.LogInformation("Call to retrieve list of team preference.");

            if (string.IsNullOrEmpty(teamId))
            {
                this.logger.LogError("TeamId is either null or empty");
                return this.BadRequest(new { message = "TeamId is either null or empty." });
            }

            try
            {
                var teamPreference = await this.teamPreferenceStorageProvider.GetTeamPreferenceAsync(teamId);
                this.RecordEvent("Team preferences - HTTP Get call succeeded");

                return this.Ok(teamPreference);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while fetching team preference details for team {teamId}.");
                throw;
            }
        }

        /// <summary>
        /// Get list of unique tags to show while configuring the preference.
        /// </summary>
        /// <param name="searchText">Search text represents the text to find and get unique tags.</param>
        /// <returns>List of unique tags.</returns>
        [HttpGet("unique-tags")]
        public async Task<IActionResult> GetUniqueTagsAsync(string searchText)
        {
            this.logger.LogInformation("Call to get list of unique tags to show while configuring the preference.");

            if (string.IsNullOrEmpty(searchText))
            {
                this.logger.LogError("Search text for GetUniqueTagsAsync is either null or empty.");
                return this.BadRequest(new { message = "Search text is either null or empty." });
            }

            try
            {
                var teamPosts = await this.teamPostSearchService.GetPostsAsync(PostSearchScope.TeamPreferenceTags, searchText, userObjectId: null);
                var uniqueTags = this.teamPreferenceStorageHelper.GetUniqueTags(teamPosts, searchText);
                this.RecordEvent("Team preferences tags - HTTP Get call succeeded");

                return this.Ok(uniqueTags);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique tags.");
                throw;
            }
        }
    }
}