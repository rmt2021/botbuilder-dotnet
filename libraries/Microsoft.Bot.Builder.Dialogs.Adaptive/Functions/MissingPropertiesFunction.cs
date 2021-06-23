// Licensed under the MIT License.
// Copyright (c) Microsoft Corporation. All rights reserved.

using AdaptiveExpressions;
using AdaptiveExpressions.Memory;
using Microsoft.Bot.Builder.Dialogs.Adaptive;
using Microsoft.Bot.Builder.Dialogs.Adaptive.Generators;

namespace Microsoft.Bot.Builder.Dialogs.Functions
{
    /// <summary>
    /// Defines missingProperties(template) expression function.
    /// </summary>
    /// <remarks>
    /// This expression will get all variables the template contains.
    /// </remarks>
    public class MissingPropertiesFunction : ExpressionEvaluator
    {
        /// <summary>
        /// Function identifier name.
        /// </summary>
        public const string Name = "missingProperties";

        private static DialogContext dialogContext;

        /// <summary>
        /// Initializes a new instance of the <see cref="MissingPropertiesFunction"/> class.
        /// </summary>
        /// <param name="context">Dialog context.</param>
        public MissingPropertiesFunction(DialogContext context)
            : base(Name, Function, ReturnType.Array, FunctionUtils.ValidateUnaryString)
        {
            dialogContext = context;
        }

        private static (object value, string error) Function(Expression expression, IMemory state, Options options)
        {
            var (args, error) = FunctionUtils.EvaluateChildren(expression, state, options);
            if (error != null)
            {
                return (null, error);
            }

            var generator = dialogContext.Services.Get<LanguageGenerator>() ?? new TemplateEngineLanguageGenerator();
            var template = args[0]?.ToString();
            var properties = generator.MissingProperties(dialogContext, template);

            return (properties, null);
        }
    }
}
