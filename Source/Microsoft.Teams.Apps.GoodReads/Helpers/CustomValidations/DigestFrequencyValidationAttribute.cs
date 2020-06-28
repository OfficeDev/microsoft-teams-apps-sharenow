// <copyright file="DigestFrequencyValidationAttribute.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoodReads.Helpers.CustomValidations
{
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Teams.Apps.GoodReads.Common;

    /// <summary>
    /// Validates digest frequency property.
    /// </summary>
    public sealed class DigestFrequencyValidationAttribute : ValidationAttribute
    {
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
                var frequency = (string)value;
                if (!string.IsNullOrEmpty(frequency))
                {
                    if (frequency == Constants.WeeklyDigest || frequency == Constants.MonthlyDigest)
                    {
                        return ValidationResult.Success;
                    }
                }
            }

            return new ValidationResult("Invalid digest frequency. Expected 'Weekly' or 'Monthly'");
        }
    }
}
