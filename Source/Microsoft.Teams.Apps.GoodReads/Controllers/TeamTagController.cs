// <copyright file="TeamTagController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GoodReads.Authentication;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Controller to handle team tags API operations.
    /// </summary>
    [Route("api/teamtag")]
    [ApiController]
    [Authorize]
    public class TeamTagController : BaseGoodReadsController
    {
        /// <summary>
        /// Used to perform logging of errors and information.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Provider to fetch, delete and upsert tags configured for team.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamTagController"/> class.
        /// </summary>
        /// <param name="logger">Used to perform logging of errors and information.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="teamTagStorageProvider">Provider to fetch, delete and upsert tags configured for team.</param>
        public TeamTagController(
            ILogger<TeamTagController> logger,
            TelemetryClient telemetryClient,
            ITeamTagStorageProvider teamTagStorageProvider)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.teamTagStorageProvider = teamTagStorageProvider;
        }

        /// <summary>
        /// Fetch configured tags for team.
        /// </summary>
        /// <param name="teamId">Team Id - unique value for each Team where tags has configured.</param>
        /// <returns>Represents Team tag entity model.</returns>
        [HttpGet]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetAsync(string teamId)
        {
            this.logger.LogInformation("Call to retrieve team tags data.");

            if (string.IsNullOrEmpty(teamId))
            {
                this.logger.LogError("Team id is either null or empty.");
                return this.BadRequest(new { message = "TeamId is either null or empty." });
            }

            try
            {
                var teamTags = await this.teamTagStorageProvider.GetTeamTagAsync(teamId);
                this.RecordEvent("Team tags - HTTP Get call succeeded");

                return this.Ok(teamTags);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team tags service.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store team tag configuration.
        /// </summary>
        /// <param name="teamTagEntity">Holds team tag detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> PostAsync([FromBody] TeamTagEntity teamTagEntity)
        {
            this.logger.LogInformation("Call to add team tag details.");

            try
            {
#pragma warning disable CA1062 // tags configuration details are validated by model validations for null check and is responded with bad request status
                var currentTeamTagConfiguration = await this.teamTagStorageProvider.GetTeamTagAsync(teamTagEntity.TeamId);
#pragma warning restore CA1062 // tags configuration details are validated by model validations for null check and is responded with bad request status

                TeamTagEntity teamTagConfiguration;

                // If there is no record in database for team, add new entry else update existing.
                if (currentTeamTagConfiguration == null)
                {
                    this.logger.LogError($"Tags configuration details were not found for team {teamTagEntity.TeamId}");
                    return this.BadRequest("Tags configuration details were not found for team");
                }
                else
                {
                    currentTeamTagConfiguration.Tags = teamTagEntity.Tags;
                    teamTagConfiguration = currentTeamTagConfiguration;
                }

                var upsertResult = await this.teamTagStorageProvider.UpsertTeamTagAsync(teamTagConfiguration);

                if (upsertResult)
                {
                    this.RecordEvent("Team tags - HTTP Post call succeeded");
                    return this.Ok(upsertResult);
                }
                else
                {
                    this.RecordEvent("Team tags - HTTP Post call failed");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, new { message = "Unable to save tags for team." });
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while saving tags for team.");
                throw;
            }
        }

        /// <summary>
        /// Get list of configured tags for a team.
        /// </summary>
        /// <param name="teamId">Team id to get the configured tags for a team.</param>
        /// <returns>List of configured tags.</returns>
        [HttpGet("configured-tags")]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetConfiguredTagsAsync(string teamId)
        {
            this.logger.LogInformation("Call to get list of configured tags for a team.");

            if (string.IsNullOrEmpty(teamId))
            {
                this.logger.LogError("TeamId is either null or empty.");
                return this.BadRequest(new { message = "TeamId is either null or empty." });
            }

            var configuredTags = new List<string>();
            try
            {
                var teamTagEntity = await this.teamTagStorageProvider.GetTeamTagAsync(teamId);
                if (teamTagEntity == null || string.IsNullOrEmpty(teamTagEntity.Tags))
                {
                    this.logger.LogInformation($"Tags are not configured for team {teamId}.");
                    return this.Ok(configuredTags);
                }

                configuredTags.AddRange(teamTagEntity.Tags.Split(';'));
                this.RecordEvent("Team tags - HTTP Get call succeeded");
                return this.Ok(configuredTags);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while fetching configured tags for team {teamId}.");
                throw;
            }
        }
    }
}