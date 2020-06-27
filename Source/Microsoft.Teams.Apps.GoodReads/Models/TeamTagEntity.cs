// <copyright file="TeamTagEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Teams.Apps.GoodReads.Helpers.CustomValidations;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// A class that represents team tag entity model.
    /// </summary>
    public class TeamTagEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets unique value for each Team where tags has configured.
        /// </summary>
        [Required]
        public string TeamId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets Azure Active Directory id of user who configured the tags in team.
        /// </summary>
        public string UserAadId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets service URL for tenant.
        /// </summary>
        [Required]
        [Url]
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets semicolon separated tags selected by user.
        /// </summary>
        [PostTagsValidation(5)]
        public string Tags { get; set; }

        /// <summary>
        /// Gets or sets date time when entry is created by user in UTC format.
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Gets or sets user name who configured tags in team.
        /// </summary>
        public string CreatedByName { get; set; }
    }
}