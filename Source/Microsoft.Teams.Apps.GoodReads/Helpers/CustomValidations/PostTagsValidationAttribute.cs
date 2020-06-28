// <copyright file="PostTagsValidationAttribute.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Helpers.CustomValidations
{
    using System.ComponentModel.DataAnnotations;

    /// <summary>
    /// Validate tag based on length and tag count for post.
    /// </summary>
    public sealed class PostTagsValidationAttribute : ValidationAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PostTagsValidationAttribute"/> class.
        /// </summary>
        /// <param name="tagsMaxCount">Max count of tags.</param>
        /// <param name="tagMaxLength">Max length of tag.</param>
        public PostTagsValidationAttribute(int tagsMaxCount, int tagMaxLength = 20)
        {
            this.TagsMaxCount = tagsMaxCount;
            this.TagMaxLength = tagMaxLength;
        }

        /// <summary>
        /// Gets max count of tags for validation.
        /// </summary>
        public int TagsMaxCount { get; }

        /// <summary>
        /// Gets max tag length for validation.
        /// </summary>
        public int TagMaxLength { get; }

        /// <summary>
        /// Validate tag based on tag length and number of tags separated by comma.
        /// </summary>
        /// <param name="value">String containing tags separated by comma.</param>
        /// <param name="validationContext">Context for getting object which needs to be validated.</param>
        /// <returns>Validation result (either error message for failed validation or success).</returns>
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            if (value != null && value.GetType() == typeof(string))
            {
                var tags = (string)value;
                if (!string.IsNullOrEmpty(tags))
                {
                    var tagsList = tags.Split(';');

                    if (tagsList.Length > this.TagsMaxCount)
                    {
                        return new ValidationResult($"Total number of tags has exceeded max count of {this.TagsMaxCount}");
                    }

                    foreach (var tag in tagsList)
                    {
                        if (string.IsNullOrWhiteSpace(tag))
                        {
                            return new ValidationResult("Tag cannot be null or empty");
                        }

                        if (tag.Length > this.TagMaxLength)
                        {
                            return new ValidationResult($"Tag length has exceeded max count of {this.TagMaxLength}");
                        }
                    }
                }
            }

            // Tags are not mandatory for adding/updating post
            return ValidationResult.Success;
        }
    }
}
