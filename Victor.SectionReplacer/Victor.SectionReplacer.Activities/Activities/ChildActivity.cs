using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Victor.SectionReplacer.Activities.Properties;
using UiPath.Shared.Activities;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.ComponentModel;
using Microsoft.Office.Interop.PowerPoint;

namespace Victor.SectionReplacer.Activities
{
	[LocalizedDisplayName(nameof(Resources.ChildActivityDisplayName))]
	[LocalizedDescription(nameof(Resources.ChildActivityDescription))]
	public class ChildActivity : AsyncTaskCodeActivity
	{
        #region Properties

        [LocalizedDisplayName(nameof(Resources.ChildActivityReplaceDisplayName))]
        [LocalizedDescription(nameof(Resources.ChildActivityReplaceDescription))]
        //[LocalizedCategory(nameof(Resources.Input))]
        [Category("Options")]
        [DefaultValue(false)]
        public bool Replace { get; set; }
        //public InArgument<Boolean> Replace { get; set; } = false;

        [LocalizedDisplayName(nameof(Resources.ChildActivityLeaveTargetOpenDisplayName))]
        [LocalizedDescription(nameof(Resources.ChildActivityLeaveTargetOpenDescription))]
        //[LocalizedCategory(nameof(Resources.Input))]
        [Category("Options")]
        [DefaultValue(true)]
        public bool LeaveTargetOpen { get; set; } = true;

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

            var targetFile = TargetFile.Get(context);
            var targetSectionName = TargetSection.Get(context);
            var srcFile = SourceFile.Get(context);
            var srcSectionName = SourceSection.Get(context);

            // TO-DO: remove hardcoded filepaths
            srcFile = @"C:\Users\Victor Weiss\Documents\UiPath\PPT Section Replacer\source.pptx";
            targetFile = @"C:\Users\Victor Weiss\Documents\UiPath\PPT Section Replacer\target.pptx";

            PPT.Application ppt = new PPT.Application();
            ppt.Visible = MsoTriState.msoTrue;
            PPT.Presentation sourcePPT = ppt.Presentations.Open(srcFile);

            // Find section in source file
            var srcSections = sourcePPT.SectionProperties;
            var srcSectionIndex = findSourceSection(srcSectionName, srcSections);
            var srcSectionStart = srcSections.FirstSlide(srcSectionIndex);
            var srcSectionEnd = srcSectionStart + srcSections.SlidesCount(srcSectionIndex) - 1;

            PPT.Presentation targetPPT = ppt.Presentations.Open(targetFile, MsoTriState.msoFalse);
            var targetSections = targetPPT.SectionProperties;
            
            var targetSectionIndex = findTargetSection(targetSectionName, targetSections);
            //var targetSectionFirstSlide = 1;
            //var targetSectionLastSlide = 1;
            var targetSectionFirstSlide = targetSections.FirstSlide(targetSectionIndex);
            var targetSectionLastSlide = targetSectionFirstSlide + targetSections.SlidesCount(targetSectionIndex) - 1;
            
            if (targetSectionFirstSlide == -1) { // meaning target section has no slides
                targetSectionFirstSlide = 0;
                for (int i = 1; i < targetSectionIndex; i++) { // find start index of section
                    targetSectionFirstSlide += targetSections.SlidesCount(i);
                }
                if (targetSectionFirstSlide == 0)
                    targetSectionFirstSlide++;
                targetSectionLastSlide = targetSectionFirstSlide;
            }

            // this can be used to preserve source formatting
            //sourcePPT.Slides[1].Copy();

            // insert slides from sourceSection into targetSection
            insertSlide(Replace, targetSections, targetSectionIndex, targetPPT, srcFile, 
                targetSectionFirstSlide, srcSectionStart, srcSectionEnd, targetSectionName, targetSectionLastSlide, ppt, targetFile);

            sourcePPT.Close();
            targetPPT.Save();
            if (!LeaveTargetOpen)
            {
                targetPPT.Close();
                ppt.Quit();
            }
            // this can be used to preserve source formatting
            //targetPPT.Windows[1].View.GotoSlide(1);
            //ppt.CommandBars.ExecuteMso("PasteSourceFormatting");

            var output = targetFile + " " + targetSectionName + " " + srcFile + " " + srcSectionName; //app.Concatenate(targetFile, targetSection, srcFile, srcSection);
            return ctx =>
            {
                SampleOutArg.Set(ctx, output); //
            };
        }

        private static void insertSlide(Boolean replace, SectionProperties targetSections, int targetSectionIndex, Presentation targetPPT, 
            string srcFile, int targetSectionFirstSlide, int srcSectionStart, int srcSectionEnd, string targetSectionName, int targetSectionLastSlide, PPT.Application ppt, string targetFile) 
        {
            try {
                if (replace) {
                    if (targetSectionIndex == 1) {
                        int targetNumSlides = targetSections.SlidesCount(1);
                        targetPPT.Slides.InsertFromFile(srcFile, targetSectionFirstSlide - 1, srcSectionStart, srcSectionEnd);
                        int srcNumSlides = srcSectionEnd - srcSectionStart + 1;
                        if (targetNumSlides > 0) {
                            for (int i = srcNumSlides + 1; i <= targetNumSlides + srcNumSlides; i++) {
                                targetPPT.Slides[srcNumSlides + 1].Delete();
                            }
                        }
                    } else {
                        targetSections.Delete(targetSectionIndex, replace);
                        targetPPT.Slides.InsertFromFile(srcFile, targetSectionFirstSlide - 1, srcSectionStart, srcSectionEnd);
                        var newSectionIndex = targetSections.AddBeforeSlide(targetSectionFirstSlide, targetSectionName);
                    }
                }
                else {
                    targetPPT.Slides.InsertFromFile(srcFile, targetSectionLastSlide, srcSectionStart, srcSectionEnd);
                }
            }
            catch (Exception e) {
                throw new Exception(e.Message + "\ntargetSectionFirstSlide: " + targetSectionFirstSlide
                    + "\ntargetSectionLastSlide: " + targetSectionLastSlide
                    + "\nsrcSectionStart: " + srcSectionStart
                    + "\nsrcSectionEnd: " + srcSectionEnd);
            }
        }

        private static int findTargetSection(string targetSectionName, SectionProperties targetSections)
        {
            var targetSectionIndex = -1;
            for (int i = 1; i <= targetSections.Count; i++) {
                if (targetSections.Name(i) == targetSectionName) {
                    targetSectionIndex = i;
                    break;
                }
            }
            if (targetSectionIndex == -1)
                throw new Exception("TargetSection " + targetSectionName + " not found in TargetFile");
            return targetSectionIndex;
        }        

        private static int findSourceSection(string srcSection, SectionProperties srcSections)
        {
            var srcSectionIndex = -1;
            for (int i = 1; i <= srcSections.Count; i++)
            {
                if (srcSections.Name(i) == srcSection)
                {
                    srcSectionIndex = i;
                    break;
                }
            }
            if (srcSectionIndex == -1)
                throw new Exception("SourceSection " + srcSection + " not found in SourceFile");
            else if (srcSections.SlidesCount(srcSectionIndex) == 0)
                throw new Exception("SourceSection " + srcSection + " has no slides");

            return srcSectionIndex;
        }

        #endregion
    }
}
