using Utilities;

namespace ProjectReport.Actions
{
    public interface IAction
    {
        public string InputFolder { get; }
        public bool Run { get; }

        public static bool ExecuteActions(IAction[] actions)
        {
            bool res = true;
            var executableActions = Array.FindAll(actions, action => action.CanExecuteAction());
            for (int i = 0; i < executableActions.Length; i++)
            {
                res = res && executableActions[i].Execute();
            }
            return res;
        }

        public static IAction[] InitiateActions(Models.Action[] userActions, string time)
        {
            List<IAction> actions = new();
            var executableUserActions = userActions.Where(ua => ua.Run).ToList();
            var actionTypes = typeof(Program).Assembly.GetTypes().Where(type => !type.IsInterface && type.IsAssignableTo(typeof(IAction))).ToList();
            foreach (var actionType in actionTypes)
            {
                SettingNameAttribute attribute = actionType.GetCustomAttributes(typeof(SettingNameAttribute), false).OfType<SettingNameAttribute>().First();
                if (executableUserActions.Select(eua => eua.Name).Contains(attribute.Name))
                {
                    actions.Add(InitializeAction(actionType, executableUserActions.Find(eua => eua.Name == attribute.Name)!, time));
                }
            }
            return actions.ToArray();
        }

        public bool CanExecuteAction()
        {
            bool res = false;
            if (Run && Directory.Exists(InputFolder))
            {
                res = true;
            }
            else
            {
                ConsoleLogger.LogError($"Directory doesn't exist: {InputFolder}", 2);
            }
            return res;
        }

        public bool Execute();

        private static IAction InitializeAction(Type actionType, Models.Action action, string time)
        {
            return actionType.Name switch
            {
                nameof(GenerateConsolidatedReportAction) => ((Func<GenerateConsolidatedReportAction>)(() =>
                {
                    var monthlyReportIdCol = action.MonthlyReportIdCol ?? 4; //TODO: default value or raise error
                    var ptrBookingMonthCol = action.PtrBookingMonthCol ?? 4;
                    var ptrProjectIdCol = action.PtrProjectIdCol ?? 4;
                    var monthlyReportMonths = action.MonthlyReportMonths ?? Array.Empty<object>();
                    var ptrBookingMonths = action.PtrBookingMonths ?? Array.Empty<object>();
                    var ptrEffortCols = action.PtrEffortCols ?? Array.Empty<object>();
                    var ptrSheetName = action.PtrSheetName ?? string.Empty;
                    return new GenerateConsolidatedReportAction(action.Run, action.InputFolder, time, monthlyReportMonths, monthlyReportIdCol, ptrBookingMonthCol, ptrBookingMonths, ptrEffortCols, ptrProjectIdCol, ptrSheetName);
                }))(),

                nameof(GenerateLeaveReportAction) => ((Func<GenerateLeaveReportAction>)(() =>
                {
                    var fy = action.FinancialYear ?? string.Empty;
                    return new GenerateLeaveReportAction(action.Run, action.InputFolder, time, fy);
                }))(),

                nameof(CalculatePunchMovementAction) => ((Func<CalculatePunchMovementAction>)(() =>
                {
                    return new CalculatePunchMovementAction(action.Run, action.InputFolder, time);
                }))(),
                _ => throw new NotImplementedException()
            };
        }
    }
}