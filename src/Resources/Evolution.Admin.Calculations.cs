{CodeGenHeader}
using BTR.Evolution.MadHatter.Web.Controls;
using BTR.Evolution.Core;
using BTR.Evolution.Data;
using Core = BTR.Evolution.MadHatter.Administration.Web.Controls.Listings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic;

namespace Administration.Controls
{
	public partial class CompareInput<T>
	{
		public BootstrapTextBox Input { get; set; }

		public bool IsInputValid(BootstrapTextBox input, DateTime? dateBirth = null)
		{
			if (string.IsNullOrEmpty(input.Text)) return false;
			if (typeof(int) == typeof(T)) return int.TryParse(input.Text, out _);
			if (typeof(double) == typeof(T)) return double.TryParse(input.Text, out _);

			if (DateTime.TryParse(input.Text, out _)) return true;

			if (dateBirth != null) return int.TryParse(input.Text, out _);

			return false;
		}

		public bool TryGetInputValue(BootstrapTextBox input, out T value, DateTime? dateBirth = null)
		{
			if (typeof(int) == typeof(T) && int.TryParse(input.Text, out var intValue))
			{
				value = intValue.ChangeType<T>();
				return true;
			}
			if (typeof(double) == typeof(T) && double.TryParse(input.Text, out var doubleValue))
			{
				value = doubleValue.ChangeType<T>();
				return true;
			}
			if (typeof(DateTime) == typeof(T))
			{
				if (DateTime.TryParse(input.Text, out var dateValue))
				{
					value = dateValue.ChangeType<T>();
					return true;
				}
				else if (dateBirth != null && int.TryParse(input.Text, out intValue))
				{
					value = dateBirth.AddYears(intValue).ChangeType<T>();
					return true;
				}
			}
			if (typeof(string) == typeof(T))
			{
				value = input.Text.ChangeType<T>();
				return true;
			}

			value = default;
			return false;
		}

		public T GetInputValue(BootstrapTextBox input, DateTime? dateBirth = null)
		{
			if (typeof(int) == typeof(T)) return int.Parse(input.Text).ChangeType<T>();
			if (typeof(double) == typeof(T)) return double.Parse(input.Text).ChangeType<T>();

			var isDate = DateTime.TryParse(input.Text, out _);

			if (isDate) return DateTime.Parse(input.Text).ChangeType<T>();

			int age = 0;
			if (dateBirth != null && int.TryParse(input.Text, out age))
			{
				return dateBirth.AddYears(age).ChangeType<T>();
			}

			throw new ArgumentOutOfRangeException("input.Text");
		}

		public bool GreaterThan(BootstrapTextBox target, DateTime? dateBirth = null) => IsInputValid(Input, dateBirth) && IsInputValid(target, dateBirth) && Comparer<T>.Default.Compare(GetInputValue(Input, dateBirth), GetInputValue(target, dateBirth)) > 0;
		public bool GreaterThanOrEqual(BootstrapTextBox target, DateTime? dateBirth = null) => IsInputValid(Input, dateBirth) && IsInputValid(target, dateBirth) && Comparer<T>.Default.Compare(GetInputValue(Input, dateBirth), GetInputValue(target, dateBirth)) >= 0;
		public bool LessThan(BootstrapTextBox target, DateTime? dateBirth = null) => IsInputValid(Input, dateBirth) && IsInputValid(target, dateBirth) && Comparer<T>.Default.Compare(GetInputValue(Input, dateBirth), GetInputValue(target, dateBirth)) < 0;
		public bool LessThanOrEqual(BootstrapTextBox target, DateTime? dateBirth = null) => IsInputValid(Input, dateBirth) && IsInputValid(target, dateBirth) && Comparer<T>.Default.Compare(GetInputValue(Input, dateBirth), GetInputValue(target, dateBirth)) <= 0;
	}

	public partial class CalculationBaseControl : EvolutionControl
	{
		protected bool IsInputValid<T>(BootstrapTextBox input) => new CompareInput<T>().IsInputValid(input);
		protected T GetInputValue<T>(BootstrapTextBox input) => new CompareInput<T>().GetInputValue(input);
		protected bool TryGetInputValue<T>(BootstrapTextBox input, out T value) => new CompareInput<T>().TryGetInputValue(input, out value);

		protected int Min(params int?[] values) => values.Any(v => v != null) ? values.Where(v => v != null).Min(v => v.Value) : int.MinValue;
		protected int Max(params int?[] values) => values.Any(v => v != null) ? values.Where(v => v != null).Max(v => v.Value) : int.MaxValue;
		protected double Min(params double?[] values) => values.Any(v => v != null) ? values.Where(v => v != null).Min(v => v.Value) : double.MinValue;
		protected double Max(params double?[] values) => values.Any(v => v != null) ? values.Where(v => v != null).Max(v => v.Value) : double.MaxValue;
		protected DateTime Min(params DateTime?[] values) => values.Any(v => v != null) ? values.Where(v => v != null).Min(v => v.Value) : DateTime.MinValue;
		protected DateTime Max(params DateTime?[] values) => values.Any(v => v != null) ? values.Where(v => v != null).Max(v => v.Value) : DateTime.MaxValue;
	}

	public partial class ProfileCalculationBaseControl : CalculationBaseControl
	{
		protected ProfileModel Profile => xDSHelper.xDSDataModel.Profile;
{HistoryAccessors}
	}
}

namespace Administration.Controls.Listings
{
	public partial class CalculationSummaryResult : Core.CalculationSummaryResult
	{
{SummaryResultTypes}
	}

	public partial class CalculationSummary : Core.CalculationSummary<CalculationSummaryResult>
	{
		protected override IEnumerable<HtmlGenericControl> GetHeaderColumns()
		{
			yield return CreateHeader("For", GetString("Labels.For"), "text-left", nameof(CalculationSummaryResult.AuthID), true);
{SummaryHeaders}
			yield return CreateHeader("Total", GetString("Labels.Total"), "text-right", nameof(CalculationSummaryResult.Total), false);
		}

		protected override IEnumerable<HtmlGenericControl> GetRowColumns(CalculationSummaryResult row)
		{
			var tag = row.IsFooter ? "th" : "td";
			yield return new HtmlGenericControl(tag, "text-left") { InnerHtml = row.AuthID ?? GetString("Labels.Total") };
{SummaryColumns}
			yield return new HtmlGenericControl(tag, "text-right table-primary total") { InnerHtml = row.Total.ToString("N0") };
		}

		protected override string[] CalculationTypes => new string[] { {CalculationTypes} };

		protected override PagingList<CalculationSummaryResult> GetCalculationListing(Core.CalculationSummaryQuery[] calculationItems, BTR.Evolution.Data.SortField sortField)
		{
			var participantSummaries =
				calculationItems
					.GroupBy(g => g.pAuthID)
					.Select(s =>
				   {
						var result =
							s.Aggregate(
								new Core.CalculationStatistics(),
								(r, cs) => r.Accumulate(cs),
								cs => cs.Compute()
							);

						return new CalculationSummaryResult
						{
							AuthID = s.Key,
{ParticipantTotals}

							Total = result.Totals.Sum(t => t.Value)
						};
					})
					.AsQueryable()
					.OrderBy(
						(sortField?.Name ?? nameof(CalculationSummaryResult.AuthID)) +
						((sortField?.Direction ?? BTR.Evolution.Data.SortDirection.Ascending) == BTR.Evolution.Data.SortDirection.Ascending
							? " asc"
							: " desc"
						)
					);

			var totals = new[] {
						new CalculationSummaryResult
						{
							IsFooter = true,
{CalculationTotals}
							Total = calculationItems.Count()
						}
					};

			var pagingListInfo = SqlPagingListInfo.GetPagingListInfo(CurrentPage, PageSize, participantSummaries.Count());

			var rows = participantSummaries.Skip((CurrentPage - 1) * PageSize).Take(PageSize).Concat(totals);

			return new PagingList<CalculationSummaryResult>
			{
				CurrentPage = pagingListInfo.CurrentPage,
				TotalPages = pagingListInfo.TotalPages,
				TotalItems = pagingListInfo.TotalItems,
				StartPage = pagingListInfo.StartPage,
				EndPage = pagingListInfo.EndPage,
				Items = rows.ToArray()
			};
		}
	}
}