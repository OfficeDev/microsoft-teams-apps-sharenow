// <copyright file="PostType.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Models
{
    /// <summary>
    /// A class that represents team tag entity model.
    /// </summary>
    public class PostType
    {
        /// <summary>
        /// Gets or sets unique value for each post type.
        /// </summary>
        public int PostTypeId { get; set; }

        /// <summary>
        /// Gets or sets post type name.
        /// </summary>
        public string PostTypeName { get; set; }

        /// <summary>
        /// Gets or sets post icon name.
        /// </summary>
        public string IconName { get; set; }
    }
}
