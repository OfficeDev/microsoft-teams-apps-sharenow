// <copyright file="TeamPreferenceStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;

    /// <summary>
    /// Implements team preference storage helper which helps to construct the model, get unique tags for team preference.
    /// </summary>
    public class TeamPreferenceStorageHelper : ITeamPreferenceStorageHelper
    {
        /// <summary>
        /// Get posts unique tags.
        /// </summary>
        /// <param name="teamPosts">Team post entities.</param>
        /// <param name="searchText">Search text for tags.</param>
        /// <returns>Represents team tags.</returns>
        public IEnumerable<string> GetUniqueTags(IEnumerable<PostEntity> teamPosts, string searchText)
        {
            if (teamPosts != null)
            {
                if (searchText == "*")
                {
                    var tagslist = new List<string>();

                    foreach (var item in teamPosts)
                    {
                        if (!string.IsNullOrEmpty(item.Tags))
                        {
                            tagslist.AddRange(item.Tags.Split(';'));
                        }
                    }

                    // Group tags based on number of occurrences and take top 50 tags having highest occurrences.
                    var filteredTags = tagslist.GroupBy(tag => tag)
                        .OrderByDescending(grouppedTags => grouppedTags.Count())
                        .Select(grouppedTags => grouppedTags.First())
                        .Take(50)
                        .OrderBy(tag => tag);
                    return filteredTags;
                }
                else
                {
                    HashSet<string> tags = new HashSet<string>();
                    foreach (var item in teamPosts)
                    {
                        if (!string.IsNullOrEmpty(item.Tags))
                        {
                            foreach (var tag in item.Tags.Split(';'))
                            {
                                tags.Add(tag);
                            }
                        }
                    }

                    return tags.Where(tag => tag.Contains(searchText, StringComparison.CurrentCulture)).OrderBy(tag => tag).Take(20);
                }
            }

            return new List<string>();
        }
    }
}