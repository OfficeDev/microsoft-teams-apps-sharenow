// <copyright file="ITeamPreferenceStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Interface for provider which helps in storing, updating team preference for posts.
    /// </summary>
    public interface ITeamPreferenceStorageProvider
    {
        /// <summary>
        /// Stores or update team preference data.
        /// </summary>
        /// <param name="teamPreferenceEntity">Holds team preference detail entity data.</param>
        /// <returns>A task that represents team preference entity data is saved or updated.</returns>
        Task<bool> UpsertTeamPreferenceAsync(TeamPreferenceEntity teamPreferenceEntity);

        /// <summary>
        /// Get team preference data.
        /// </summary>
        /// <param name="teamId">Team Id for which need to fetch data.</param>
        /// <returns>A task that represents to hold team preference data.</returns>
        Task<TeamPreferenceEntity> GetTeamPreferenceAsync(string teamId);

        /// <summary>
        /// Get team preferences data.
        /// </summary>
        /// <param name="digestFrequency">Digest frequency text for notification like Monthly/Weekly.</param>
        /// <returns>A task that represent collection to hold team preferences data.</returns>
        Task<IEnumerable<TeamPreferenceEntity>> GetTeamPreferencesByDigestFrequencyAsync(string digestFrequency);

        /// <summary>
        /// Delete team preference if Bot is uninstalled.
        /// </summary>
        /// <param name="teamId">Holds team id.</param>
        /// <returns>A boolean that represents team tags is successfully deleted or not.</returns>
        Task<bool> DeleteTeamPreferenceAsync(string teamId);
    }
}
