using System;
using System.Activities;
using System.ComponentModel;
using System.Activities.Statements;
using System.Threading;
using System.Threading.Tasks;
using Victor.SectionReplacer.Activities.Properties;
using Victor.SectionReplacer;
using UiPath.Shared.Activities;

namespace Victor.SectionReplacer.Activities
{

    [LocalizedDescription(nameof(Resources.ParentScopeDescription))]
    [LocalizedDisplayName(nameof(Resources.ParentScope))]
    public class ParentScope : AsyncTaskNativeActivity
    {
		#region Properties
    
        private Application Application;

		public const string ApplicationTag = "ParentScopeApplication";

		[Browsable(false)]
        public ActivityAction<Application> Body { get; set; }

        #endregion


        #region Constructors

        public ParentScope()
        {
            Body = new ActivityAction<Application>
            {
                Argument = new DelegateInArgument<Application>(ApplicationTag),
                Handler = new Sequence { DisplayName = "Do" }
            };
        }

        #endregion


        #region Private Methods

        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<NativeActivityContext>> ExecuteAsync(NativeActivityContext context, CancellationToken cancellationToken)
        {
            // Schedule the child activities in the scope's body to run and make the client available to them
            if (Body != null)
				context.ScheduleAction<Application > (Body, Application, OnCompleted, OnFaulted);

			// Any actions to perform after the child activities are scheduled go here
			return _ => { };
		}

        private void OnFaulted(NativeActivityFaultContext faultContext, Exception propagatedException, ActivityInstance propagatedFrom)
        {
        }

        private void OnCompleted(NativeActivityContext context, ActivityInstance completedInstance)
        {
        }

        #endregion


        #region Helpers
        
        #endregion
    }
}
