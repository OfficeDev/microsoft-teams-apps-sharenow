// <copyright file="UserTeamMembershipEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// A class that represents user team membership entity.
    /// </summary>
    public class UserTeamMembershipEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets Azure Active Directory id of user.
        /// </summary>
        public string UserAadObjectId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets id of the team.
        /// </summary>
        public string TeamId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets service URL where responses to this activity should be sent.
        /// </summary>
        public string ServiceUrl { get; set; }
    }
}
