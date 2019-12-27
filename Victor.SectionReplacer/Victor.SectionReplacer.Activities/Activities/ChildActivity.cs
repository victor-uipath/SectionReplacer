using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Victor.SectionReplacer.Activities.Properties;
using UiPath.Shared.Activities;

namespace Victor.SectionReplacer.Activities
{
	[LocalizedDisplayName(nameof(Resources.ChildActivityDisplayName))]
	[LocalizedDescription(nameof(Resources.ChildActivityDescription))]
	public class ChildActivity : AsyncTaskCodeActivity
	{
		#region Properties

		[LocalizedDisplayName(nameof(Resources.ChildActivityReplaceDisplayName))]
		[LocalizedDescription(nameof(Resources.ChildActivityReplaceDescription))]
		[LocalizedCategory(nameof(Resources.Input))]
		public InArgument<Boolean> Replace { get; set; }

        [LocalizedDisplayName(nameof(Resources.ChildActivitySourceFileDisplayName))]
        [LocalizedDescription(nameof(Resources.ChildActivitySourceFileDescription))]
        [LocalizedCategory(nameof(Resources.Input))]
        public InArgument<String> SourceFile { get; set; }

        [LocalizedDisplayName(nameof(Resources.ChildActivitySourceSectionDisplayName))]
        [LocalizedDescription(nameof(Resources.ChildActivitySourceSectionDescription))]
        [LocalizedCategory(nameof(Resources.Input))]
        public InArgument<String> SourceSection { get; set; }

        [LocalizedDisplayName(nameof(Resources.ChildActivityTargetFileDisplayName))]
        [LocalizedDescription(nameof(Resources.ChildActivityTargetFileDescription))]
        [LocalizedCategory(nameof(Resources.Input))]
        public InArgument<String> TargetFile { get; set; }

        [LocalizedDisplayName(nameof(Resources.ChildActivityTargetSectionDisplayName))]
        [LocalizedDescription(nameof(Resources.ChildActivityTargetSectionDescription))]
        [LocalizedCategory(nameof(Resources.Input))]
        public InArgument<String> TargetSection { get; set; }

        [LocalizedDisplayName(nameof(Resources.ChildActivitySumDisplayName))]
		[LocalizedDescription(nameof(Resources.ChildActivitySumDescription))]
		[LocalizedCategory(nameof(Resources.Output))]
		public OutArgument<String> SampleOutArg { get; set; }
        
        #endregion

        public ChildActivity()
        {
            Constraints.Add(ActivityConstraints.HasParentType<ChildActivity, ParentScope>(Resources.ValidationMessage));
        }

        #region Protected Methods

        /// <summary>
        /// Validates properties at design-time.
        /// </summary>
        /// <param name="metadata"></param>
        protected override void CacheMetadata(CodeActivityMetadata metadata)
		{
			if (SourceFile == null) metadata.AddValidationError(string.Format(Resources.MetadataValidationError, nameof(SourceFile)));
            if (SourceSection == null) metadata.AddValidationError(string.Format(Resources.MetadataValidationError, nameof(SourceSection)));
            if (TargetFile == null) metadata.AddValidationError(string.Format(Resources.MetadataValidationError, nameof(TargetFile)));
            if (TargetSection == null) metadata.AddValidationError(string.Format(Resources.MetadataValidationError, nameof(TargetSection)));

            base.CacheMetadata(metadata);
    }

    /// <summary>
    /// Runs the main logic of the activity. Has access to the context, 
    /// which holds the values of properties for this activity and those from the parent scope.
    /// </summary>
    /// <param name="context"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
		{
            var property = context.DataContext.GetProperties()[ParentScope.ApplicationTag];
            var app = property.GetValue(context.DataContext) as Application;
           
            var replace = Replace.Get(context);
            var targetFile = TargetFile.Get(context);
            var targetSection = TargetSection.Get(context);
            var sourceFile = SourceFile.Get(context);
            var sourceSection = SourceSection.Get(context);

            var output = targetFile + " " + targetSection + " " + sourceFile + " " + sourceSection; //app.Concatenate(targetFile, targetSection, sourceFile, sourceSection);
            return ctx =>
            {
                SampleOutArg.Set(ctx, output);
            };
        }

        #endregion
    }
}
