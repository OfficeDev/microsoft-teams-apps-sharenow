// <copyright file="PostTypeHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Helpers
{
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    ///  A class that handles the post types based on the post type id.
    /// </summary>
    public class PostTypeHelper
    {
        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="PostTypeHelper"/> class.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public PostTypeHelper(IStringLocalizer<Strings> localizer)
        {
            this.localizer = localizer;
        }

        /// <summary>
        /// Valid post types.
        /// </summary>
        public enum PostTypeEnum
        {
            /// <summary>
            /// No post.
            /// </summary>
            None = 0,

            /// <summary>
            /// Blog post type.
            /// </summary>
            BlogPost = 1,

            /// <summary>
            /// Other post type.
            /// </summary>
            Other = 2,

            /// <summary>
            /// Podcast post type.
            /// </summary>
            Podcast = 3,

            /// <summary>
            /// Video post type.
            /// </summary>
            Video = 4,

            /// <summary>
            /// Book post type.
            /// </summary>
            Book = 5,
        }

        /// <summary>
        /// Get the post type using its id.
        /// </summary>
        /// <param name="key">Post type id value.</param>
        /// <returns>Returns a post type from the id value.</returns>
        public PostType GetPostType(int key)
        {
            return key switch
            {
                (int)PostTypeEnum.BlogPost =>
                    new PostType { PostTypeName = this.localizer.GetString("BlogPostType"), IconName = "blogTypeDot.png", PostTypeId = 1 },

                (int)PostTypeEnum.Other =>
                    new PostType { PostTypeName = this.localizer.GetString("OtherPostType"), IconName = "otherTypeDot.png", PostTypeId = 2 },

                (int)PostTypeEnum.Podcast =>
                    new PostType { PostTypeName = this.localizer.GetString("PodcastPostType"), IconName = "podcastTypeDot.png", PostTypeId = 3 },

                (int)PostTypeEnum.Video =>
                    new PostType { PostTypeName = this.localizer.GetString("VideoPostType"), IconName = "videoTypeDot.png", PostTypeId = 4 },

                (int)PostTypeEnum.Book =>
                    new PostType { PostTypeName = this.localizer.GetString("BookPostType"), IconName = "bookTypeDot.png", PostTypeId = 5 },

                _ => null,
            };
        }
    }
}
