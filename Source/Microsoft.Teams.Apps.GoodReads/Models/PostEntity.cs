// <copyright file="PostEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Azure.Search;
    using Microsoft.Teams.Apps.GoodReads.Helpers;
    using Microsoft.Teams.Apps.GoodReads.Helpers.CustomValidations;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// A class that represents post entity model which helps to create, insert, update and delete the post.
    /// </summary>
    public class PostEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets Azure Active Directory id of author who created the post.
        /// </summary>
        [IsFilterable]
        public string UserId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets unique identifier for each created post.
        /// </summary>
        [Key]
        [IsFilterable]
        public string PostId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets user selected value (type of post) from the dropdown list. For valid values check <see cref="PostTypeHelper"/>.
        /// </summary>
        [IsFilterable]
        [Required]
        [Range(1, 5)]
        public int Type { get; set; }

        /// <summary>
        /// Gets or sets title of post.
        /// </summary>
        [IsSearchable]
        [Required]
        [MaxLength(100)]
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets user entered post description value.
        /// </summary>
        [Required]
        [MinLength(150)]
        [MaxLength(300)]
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets URL of the content (article).
        /// </summary>
        [Required]
        [Url]
        [MaxLength(400)]
        public string ContentUrl { get; set; }

        /// <summary>
        /// Gets or sets semicolon separated tags entered by user.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [PostTagsValidation(3)]
        public string Tags { get; set; }

        /// <summary>
        /// Gets or sets date time when entry is created.
        /// </summary>
        [IsSortable]
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Gets or sets author name who created post.
        /// </summary>
        [IsFilterable]
        public string CreatedByName { get; set; }

        /// <summary>
        /// Gets or sets date time when entry is updated.
        /// </summary>
        [IsSortable]
        public DateTime UpdatedDate { get; set; }

        /// <summary>
        /// Gets or sets total number of likes received for a post by users.
        /// </summary>
        [IsSortable]
        public int TotalVotes { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the post is deleted by user.
        /// </summary>
        [IsFilterable]
        public bool IsRemoved { get; set; }
    }
}
