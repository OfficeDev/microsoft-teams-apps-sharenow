// <copyright file="GoodReadsActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.GoodReads.Models.Configuration
{
    /// <summary>
    /// This class provide options for the <see cref="GoodReadsActivityHandlerOptions" /> bot.
    /// </summary>
    public sealed class GoodReadsActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets application base URL used to return success or failure task module result.
        /// </summary>
        public string AppBaseUri { get; set; }

        /// <summary>
        /// Gets or sets entity id of static discover tab.
        /// </summary>
        public string DiscoverTabEntityId { get; set; }
    }
}
