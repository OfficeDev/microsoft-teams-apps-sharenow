// <copyright file="MustBeTeamMemberUserPolicyHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Authentication
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc.Filters;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Teams.Apps.GoodReads.Common;
    using Microsoft.Teams.Apps.GoodReads.Common.Interfaces;
    using Microsoft.Teams.Apps.GoodReads.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// This class is an authorization handler, which handles the authorization requirement.
    /// </summary>
    public class MustBeTeamMemberUserPolicyHandler : AuthorizationHandler<MustBeValidTeamMemberRequirement>
    {
        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter botAdapter;

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Provider to fetch team configuration for tab.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// Cache for storing authorization result.
        /// </summary>
        private readonly IMemoryCache memoryCache;

        /// <summary>
        /// Initializes a new instance of the <see cref="MustBeTeamMemberUserPolicyHandler"/> class.
        /// </summary>
        /// <param name="botAdapter">Bot adapter for getting team members.</param>
        /// <param name="microsoftAppCredentials">Microsoft application credentials.</param>
        /// <param name="teamTagStorageProvider">Provider to fetch team configuration for tab.</param>
        /// <param name="memoryCache">MemoryCache instance for caching authorization result.</param>
        public MustBeTeamMemberUserPolicyHandler(IBotFrameworkHttpAdapter botAdapter, MicrosoftAppCredentials microsoftAppCredentials, ITeamTagStorageProvider teamTagStorageProvider, IMemoryCache memoryCache)
        {
            this.teamTagStorageProvider = teamTagStorageProvider;
            this.memoryCache = memoryCache;
            this.botAdapter = botAdapter;
            this.microsoftAppCredentials = microsoftAppCredentials;
        }

        /// <summary>
        /// This method handles the authorization requirement.
        /// </summary>
        /// <param name="context">AuthorizationHandlerContext instance.</param>
        /// <param name="requirement">IAuthorizationRequirement instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override Task HandleRequirementAsync(
            AuthorizationHandlerContext context,
            MustBeValidTeamMemberRequirement requirement)
        {
            context = context ?? throw new ArgumentNullException(nameof(context));

            string teamId = string.Empty;
            var oidClaimType = "http://schemas.microsoft.com/identity/claims/objectidentifier";

            var claim = context.User.Claims.FirstOrDefault(p => oidClaimType.Equals(p.Type, StringComparison.OrdinalIgnoreCase));

            if (context.Resource is AuthorizationFilterContext authorizationFilterContext)
            {
                // Wrap the request stream so that we can rewind it back to the start for regular request processing.
                authorizationFilterContext.HttpContext.Request.EnableBuffering();

                if (string.IsNullOrEmpty(authorizationFilterContext.HttpContext.Request.QueryString.Value))
                {
                    // Read the request body, parse out the team tag entity object to get team Id.
                    var streamReader = new StreamReader(authorizationFilterContext.HttpContext.Request.Body, Encoding.UTF8, true, 1024, leaveOpen: true);
                    using var jsonReader = new JsonTextReader(streamReader);
                    var obj = JObject.Load(jsonReader);
                    var tagEntity = obj.ToObject<TeamTagEntity>();
                    authorizationFilterContext.HttpContext.Request.Body.Seek(0, SeekOrigin.Begin);
                    teamId = tagEntity.TeamId;
                }
                else
                {
                    var requestQuery = authorizationFilterContext.HttpContext.Request.Query;
                    teamId = requestQuery.Where(queryData => queryData.Key == "teamId").Select(queryData => queryData.Value.ToString()).FirstOrDefault();
                }
            }

            if (this.ValidateUserAsync(teamId, claim.Value).Result)
            {
                context.Succeed(requirement);
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Check if a user is a member of a certain team.
        /// </summary>
        /// <param name="teamId">The team id that the validator uses to check if the user is a member of the team. </param>
        /// <param name="userAadObjectId">The user's Azure Active Directory object id.</param>
        /// <returns>The flag indicates that the user is a part of certain team or not.</returns>
        private async Task<bool> ValidateUserAsync(string teamId, string userAadObjectId)
        {
            this.memoryCache.TryGetValue(userAadObjectId, out bool isUserValid);
            if (isUserValid == false)
            {
                var userTeamMembershipEntities = await this.teamTagStorageProvider.GetTeamTagAsync(teamId);

                if (userTeamMembershipEntities == null)
                {
                    return false;
                }

                TeamsChannelAccount teamMember = new TeamsChannelAccount();

                var conversationReference = new ConversationReference
                {
                    ChannelId = Constants.TeamsBotFrameworkChannelId,
                    ServiceUrl = userTeamMembershipEntities.ServiceUrl,
                };
                await ((BotFrameworkAdapter)this.botAdapter).ContinueConversationAsync(
                    this.microsoftAppCredentials.MicrosoftAppId,
                    conversationReference,
                    async (context, token) =>
                    {
                        teamMember = await TeamsInfo.GetTeamMemberAsync(context, userAadObjectId, teamId, CancellationToken.None);
                    }, default);

                var isValid = teamMember != null;
                this.memoryCache.Set(userAadObjectId, isValid, TimeSpan.FromHours(1));
                return isValid;
            }

            return isUserValid;
        }
    }
}
