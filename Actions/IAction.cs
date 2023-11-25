﻿using Services;
using Utilities;

namespace Actions
{
    public abstract class Action
    {

        public static string Time { get; private set; }
        public static ILogger Logger { get; private set; }
        public static DataService DataService { get; private set; }
        public static ReadService ReadService { get; private set; }
        public static WriteService WriteService { get; private set; }
        public static ExportService ExportService { get; private set; }

        public static void Init(string time, ILogger logger, DataService dataService, ReadService readService, WriteService writeService, ExportService exportService)
        {
            Time = time;
            Logger = logger;
            DataService = dataService;
            ReadService = readService;
            WriteService = writeService;
            ExportService = exportService;
        }

        public static bool ExecuteActions(IEnumerable<Models.Action> userActions, CancellationToken cancellationToken)
        {
            bool res = true;
            var executableActions = userActions.Where(ua => ua.Run);
            var actions = InitiateActions(executableActions);
            foreach (var action in actions)
            {
                if (cancellationToken.IsCancellationRequested) return false;
                res = res && action.ValidateAndRun();
            }
            return res;
        }

        public abstract bool Validate();

        public abstract bool Run();

        public bool ValidateAndRun()
        {
            if (Validate())
            {
                return Run();
            }
            return false;
        }

        private static Action InitializeAction(Type actionType, Models.Action action)
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
                    return new GenerateConsolidatedReportAction(
                        action.InputFolder,
                        monthlyReportMonths,
                        monthlyReportIdCol,
                        ptrBookingMonthCol,
                        ptrBookingMonths,
                        ptrEffortCols,
                        ptrProjectIdCol,
                        ptrSheetName);
                }))(),

                nameof(GenerateLeaveReportAction) => ((Func<GenerateLeaveReportAction>)(() => new GenerateLeaveReportAction(action.InputFolder, action.FinancialYear ?? string.Empty)))(),

                nameof(CalculatePunchMovementAction) => ((Func<CalculatePunchMovementAction>)(() => new CalculatePunchMovementAction(action.InputFolder, action.CutOff ?? string.Empty)))(),

                nameof(InOutEntryAction) => ((Func<InOutEntryAction>)(() => new InOutEntryAction(action.InputFolder)))(),

                _ => throw new NotImplementedException("Action not implemented.")
            };
        }

        private static IEnumerable<Action> InitiateActions(IEnumerable<Models.Action> userActions)
        {
            IEnumerable<Action> actions = Enumerable.Empty<Action>();
            var executableUserActions = userActions.Where(ua => ua.Run).ToList();
            var actionTypes = typeof(Action).Assembly.GetTypes().Where(type => type != typeof(Action) && type.IsAssignableTo(typeof(Action))).ToList();
            foreach (var actionType in actionTypes)
            {
                ActionNameAttribute attribute = actionType.GetCustomAttributes(typeof(ActionNameAttribute), false).OfType<ActionNameAttribute>().First();
                if (executableUserActions.Select(eua => eua.Name).Contains(attribute.Name))
                {
                    actions =  actions.Append(InitializeAction(actionType, executableUserActions.Find(eua => eua.Name == attribute.Name)!));
                }
            }
            return actions;
        }
    }
}