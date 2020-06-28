// <copyright file="ITeamTagStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Interface for provider which helps in storing, updating or deleting team tags.
    /// </summary>
    public interface ITeamTagStorageProvider
    {
        /// <summary>
        /// Stores or update team tags data.
        /// </summary>
        /// <param name="teamTagEntity">Holds team preference detail entity data.</param>
        /// <returns>A task that represents team preference entity data is saved or updated.</returns>
        Task<bool> UpsertTeamTagAsync(TeamTagEntity teamTagEntity);

        /// <summary>
        /// Get team tags data.
        /// </summary>
        /// <param name="teamId">Team id for which need to fetch data.</param>
        /// <returns>A task that represents to hold team tags data.</returns>
        Task<TeamTagEntity> GetTeamTagAsync(string teamId);

        /// <summary>
        /// Delete configured tags for a team if Bot is uninstalled.
        /// </summary>
        /// <param name="teamId">Holds team id.</param>
        /// <returns>A task that represents team tags data is deleted.</returns>
        Task<bool> DeleteTeamTagAsync(string teamId);
    }
}
