using Services;
using Utilities;

namespace Actions
{
    public abstract class Action : IAction
    {
        public static DataService DataService { get; set; } = new(Logger!);

        public static string ExportFolder => @$"{InputFolder}\Reports_{Time}";
        public static ExportService ExportService { get; set; } = new(Logger!);

        public static string InputFolder { get; set; } = string.Empty;
        public static ILogger Logger { get; set; } = new ConsoleLogger();

        public static ReadService ReadService { get; set; } = new(Logger!);

        public static string Time { get; set; } = string.Empty;
        public static WriteService WriteService { get; set; } = new(Logger!);

        public static bool ExecuteActions(IEnumerable<Models.Action> userActions, CancellationToken cancellationToken)
        {
            bool res = true;
            var executableActions = userActions.Where(ua => ua.Run);
            var actions = InitiateActions(executableActions);
            foreach (var action in actions)
            {
                if (cancellationToken.IsCancellationRequested) return false;
                res = res && action.InitValidateAndRun(GetActionName(action));
            }
            return res;
        }

        public static void Initialize(string time, ILogger logger, DataService dataService, ReadService readService, WriteService writeService, ExportService exportService)
        {
            (Time, Logger, DataService, ReadService, WriteService, ExportService) = (time, logger, dataService, readService, writeService, exportService);
        }

        public static bool ValidateDirectory(string directory)
        {
            if (!Directory.Exists(directory))
            {
                Logger.LogError($"Directory doesn't exist: {directory}.", 2);
                return false;
            }
            return true;
        }

        public static bool ValidateReports(List<string> reports, string errorMessage)
        {
            if (reports.Count == 0)
            {
                Logger.LogError(errorMessage);
                return false;
            }
            return true;
        }

        public abstract void Init();

        public abstract bool Run();

        public abstract bool Validate();

        private static string? GetActionName(Action action) => action.ToString()?.Replace("Actions.", "").Replace("Action", "");

        private static IEnumerable<Action> InitiateActions(IEnumerable<Models.Action> userActions)
        {
            Action InitializeAction(Type actionType, Models.Action action)
            {
                InputFolder = action.InputFolder;
                return actionType.Name switch
                {
                    nameof(GenerateConsolidatedReportAction) => ((Func<GenerateConsolidatedReportAction>)(() => new GenerateConsolidatedReportAction(action.MonthlyReportMonths,
                                                                                                                                                     action.MonthlyReportIdCol,
                                                                                                                                                     action.PtrBookingMonthCol,
                                                                                                                                                     action.PtrBookingMonths,
                                                                                                                                                     action.PtrEffortCols,
                                                                                                                                                     action.PtrProjectIdCol,
                                                                                                                                                     action.PtrSheetName)))(),

                    nameof(GenerateLeaveReportAction) => ((Func<GenerateLeaveReportAction>)(() => new GenerateLeaveReportAction(action.FinancialYear)))(),

                    nameof(CalculatePunchMovementAction) => ((Func<CalculatePunchMovementAction>)(() => new CalculatePunchMovementAction(action.CutOff)))(),

                    nameof(MonthlyReportInOutEntryAction) => ((Func<MonthlyReportInOutEntryAction>)(() => new MonthlyReportInOutEntryAction()))(),

                    nameof(AttendanceReportEntryAction) => ((Func<AttendanceReportEntryAction>)(() => new AttendanceReportEntryAction()))(),

                    _ => throw new NotImplementedException("Action not implemented.")
                };
            }

            IEnumerable<Action> actions = [];
            var executableUserActions = userActions.Where(ua => ua.Run).ToList();
            var actionTypes = typeof(Action).Assembly.GetTypes().Where(type => type != typeof(Action) && type.IsAssignableTo(typeof(Action))).ToList();
            foreach (var actionType in actionTypes)
            {
                ActionNameAttribute attribute = actionType.GetCustomAttributes(typeof(ActionNameAttribute), false).OfType<ActionNameAttribute>().First();
                if (executableUserActions.Select(eua => eua.Name).Contains(attribute.Name))
                {
                    actions = actions.Append(InitializeAction(actionType, executableUserActions.Find(eua => eua.Name == attribute.Name)!));
                }
            }
            return actions;
        }

        private bool InitValidateAndRun(string? actionName)
        {
            if (!string.IsNullOrEmpty(actionName))
            {
                Logger.LogInfo($"Running action {actionName}:", 1);
            }
            Logger.LogChar('-', 100);
            Logger.LogLine(1);
            Init();
            return Validate() && Run();
        }
    }
}