// <copyright file="AdaptiveTaskModuleCardAction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Models.Card
{
    using Newtonsoft.Json;

    /// <summary>
    /// A class that represents the adaptive submit action model.
    /// </summary>
    public class AdaptiveTaskModuleCardAction
    {
        /// <summary>
        /// Gets or sets action type for button.
        /// </summary>
        [JsonProperty("type")]
        public string Type
        {
            get
            {
                return "task/fetch";
            }
            set => this.Type = "task/fetch";
        }

        /// <summary>
        /// Gets or sets bot command to be used by bot for processing user inputs.
        /// </summary>
        [JsonProperty("text")]
        public string Text { get; set; }
    }
}
