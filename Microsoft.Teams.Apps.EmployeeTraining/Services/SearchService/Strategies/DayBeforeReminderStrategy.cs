// <copyright file="DayBeforeReminderStrategy.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.EmployeeTraining.Services.SearchService.Strategies
{
    using System;
    using System.Globalization;
    using Microsoft.Teams.Apps.EmployeeTraining.Models;

    /// <summary>
    /// Generates filter query for fetching events to send day before notifications.
    /// </summary>
    public class DayBeforeReminderStrategy : IFilterGeneratingStrategy
    {
        /// <inheritdoc/>
        public string GenerateFilterQuery(SearchParametersDto searchParametersDto)
        {
            var startDate = DateTime.UtcNow.Date.AddDays(1).Date;
            var endDate = DateTime.UtcNow.Date.AddDays(2).Date;

            return $"{nameof(EventEntity.Status)} eq {(int)EventStatus.Active} and " +
                $"{nameof(EventEntity.StartDate)} ge {startDate.ToString("O", CultureInfo.InvariantCulture)} and " +
                $"{nameof(EventEntity.StartDate)} le {endDate.ToString("O", CultureInfo.InvariantCulture)} and " +
                $"{nameof(EventEntity.RegisteredAttendeesCount)} gt 0";
        }
    }
}