// <copyright file="UserPrivatePostEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// A class that represents user private post model.
    /// </summary>
    public class UserPrivatePostEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets unique Azure Active Directory Id of user.
        /// </summary>
        public string UserId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets unique identifier for each created post.
        /// </summary>
        [Required]
        public string PostId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets date time when entry is created.
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Gets or sets name of user who created the post.
        /// </summary>
        public string CreatedByName { get; set; }
    }
}
