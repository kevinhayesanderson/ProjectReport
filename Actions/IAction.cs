using Services;
using Utilities;

namespace Actions
{
    public interface IAction
    {
        public string InputFolder { get; }
        public bool Run { get; }

        public static bool ExecuteActions(Models.Action[] userActions, string time, ILogger logger, DataService dataService, ReadService readService, WriteService writeService, ExportService exportService, CancellationToken cancellationToken)
        {
            bool res = true;
            var actions = InitiateActions(userActions, time, logger, dataService, readService, writeService, exportService);
            var executableActions = Array.FindAll(actions, action => action.CanExecuteAction(logger));
            for (int i = 0; i < executableActions.Length; i++)
            {
                if (cancellationToken.IsCancellationRequested) return false;
                res = res && executableActions[i].Execute();
            }
            return res;
        }

        private static IAction[] InitiateActions(Models.Action[] userActions, string time, ILogger logger, DataService dataService, ReadService readService, WriteService writeService, ExportService exportService)
        {
            List<IAction> actions = [];
            var executableUserActions = userActions.Where(ua => ua.Run).ToList();
            var actionTypes = typeof(IAction).Assembly.GetTypes().Where(type => !type.IsInterface && type.IsAssignableTo(typeof(IAction))).ToList();
            foreach (var actionType in actionTypes)
            {
                ActionNameAttribute attribute = actionType.GetCustomAttributes(typeof(ActionNameAttribute), false).OfType<ActionNameAttribute>().First();
                if (executableUserActions.Select(eua => eua.Name).Contains(attribute.Name))
                {
                    actions.Add(InitializeAction(actionType, executableUserActions.Find(eua => eua.Name == attribute.Name)!, time, logger, dataService, readService, writeService, exportService));
                }
            }
            return [.. actions];
        }

        public bool CanExecuteAction(ILogger logger)
        {
            bool res = false;
            if (Run && Directory.Exists(InputFolder))
            {
                res = true;
            }
            else
            {
                logger.LogError($"Directory doesn't exist: {InputFolder}", 2);
            }
            return res;
        }

        public bool Execute();

        private static IAction InitializeAction(Type actionType, Models.Action action, string time, ILogger logger, DataService dataService, ReadService readService, WriteService writeService, ExportService exportService)
        {
            return actionType.Name switch
            {
                nameof(GenerateConsolidatedReportAction) => ((Func<GenerateConsolidatedReportAction>)(() =>
                {
                    var monthlyReportIdCol = action.MonthlyReportIdCol ?? 4; //TODO: default value or raise error
                    var ptrBookingMonthCol = action.PtrBookingMonthCol ?? 4;
                    var ptrProjectIdCol = action.PtrProjectIdCol ?? 4;
                    var monthlyReportMonths = action.MonthlyReportMonths ?? [];
                    var ptrBookingMonths = action.PtrBookingMonths ?? [];
                    var ptrEffortCols = action.PtrEffortCols ?? [];
                    var ptrSheetName = action.PtrSheetName ?? string.Empty;
                    return new GenerateConsolidatedReportAction(action.Run, action.InputFolder, time, logger, monthlyReportMonths, monthlyReportIdCol, ptrBookingMonthCol, ptrBookingMonths, ptrEffortCols, ptrProjectIdCol, ptrSheetName, dataService, readService, exportService);
                }))(),

                nameof(GenerateLeaveReportAction) => ((Func<GenerateLeaveReportAction>)(() =>
                {
                    var fy = action.FinancialYear ?? string.Empty;
                    return new GenerateLeaveReportAction(action.Run, action.InputFolder, time, logger, fy, exportService);
                }))(),

                nameof(CalculatePunchMovementAction) => ((Func<CalculatePunchMovementAction>)(() =>
                {
                    return new CalculatePunchMovementAction(action.Run, action.InputFolder, time, logger);
                }))(),
                _ => throw new NotImplementedException("Action not implemented.")
            };
        }
    }
}