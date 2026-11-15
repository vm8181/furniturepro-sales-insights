"""Semantic model for the FurniturePro report: Power Query (M) tables and DAX measures, written as TMDL."""
import json
import uuid
from pathlib import Path

SOURCE_DEFAULT = r"C:\Projects\furniturepro-sales-insights\data\FurniturePro_Data.xlsx"


def tag(*parts) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "furniturepro/" + "/".join(parts)))


def q(name: str) -> str:
    """TMDL name: quoted when it has anything but letters, digits and underscore."""
    return name if name.replace("_", "").isalnum() else "'" + name + "'"


# ---------------------------------------------------------------------------- shared queries (not loaded)

EXPRESSIONS = []  # name, doc, M, result type

EXPRESSIONS.append(("Workbook", "The Excel workbook, opened once and reused by every query.", """let
    Source = Excel.Workbook(File.Contents(SourceFile), null, true)
in
    Source""", "Table"))

EXPRESSIONS.append(("OrdersBase", "Cleaned order lines with every added column. The Orders, Customer, State, Date and Top N tables all start from this.", """let
    Sheet = Workbook{[Item = "Orders", Kind = "Sheet"]}[Data],
    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    // The sheet's Year, Country and Postal Code are not needed: the year comes from Order Date
    Kept = Table.SelectColumns(Headers, {"Row ID", "Order ID", "Order Date", "Ship Mode", "Customer ID", "Segment", "City", "State", "Region", "Product ID", "Sales Amount", "Quantity", "Discount", "Profit", "SalesPerson"}),
    Typed = Table.TransformColumnTypes(Kept, {{"Row ID", Int64.Type}, {"Order ID", type text}, {"Order Date", type date}, {"Ship Mode", type text}, {"Customer ID", type text}, {"Segment", type text}, {"City", type text}, {"State", type text}, {"Region", type text}, {"Product ID", type text}, {"Sales Amount", type number}, {"Quantity", Int64.Type}, {"Discount", type number}, {"Profit", type number}, {"SalesPerson", type text}}, "en-US"),
    Trimmed = Table.TransformColumns(Typed, {{"Order ID", Text.Trim, type text}, {"Ship Mode", Text.Trim, type text}, {"Customer ID", Text.Trim, type text}, {"Segment", Text.Trim, type text}, {"City", Text.Trim, type text}, {"State", Text.Trim, type text}, {"Region", Text.Trim, type text}, {"Product ID", Text.Trim, type text}, {"SalesPerson", Text.Trim, type text}}),
    Valid = Table.SelectRows(Trimmed, each [Order ID] <> null and [Order Date] <> null and [Sales Amount] <> null),
    Renamed = Table.RenameColumns(Valid, {{"SalesPerson", "Sales Person ID"}}),
    // Returns are listed per order, so a returned order marks all of its lines as returned
    ReturnsSheet = Table.PromoteHeaders(Workbook{[Item = "Returns", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
    ReturnedOnly = Table.SelectRows(ReturnsSheet, each [Order ID] <> null and Text.Lower(Text.Trim(Text.From([Returned]))) = "yes"),
    ReturnedOrders = Table.Distinct(Table.TransformColumns(Table.SelectColumns(ReturnedOnly, {"Order ID"}), {{"Order ID", each Text.Trim(Text.From(_)), type text}})),
    Merged = Table.NestedJoin(Renamed, {"Order ID"}, ReturnedOrders, {"Order ID"}, "ReturnMatch", JoinKind.LeftOuter),
    AddReturned = Table.AddColumn(Merged, "Returned", each if Table.IsEmpty([ReturnMatch]) then "No" else "Yes", type text),
    RemovedMatch = Table.RemoveColumns(AddReturned, {"ReturnMatch"}),
    AddReturnedAmount = Table.AddColumn(RemovedMatch, "Returned Amount", each if [Returned] = "Yes" then [Sales Amount] else 0, type number),
    AddNetSales = Table.AddColumn(AddReturnedAmount, "Net Sales", each [Sales Amount] - [Returned Amount], type number),
    // List Price is the value before discount: Sales Amount = List Price x (1 - Discount)
    AddListPrice = Table.AddColumn(AddNetSales, "List Price", each if [Discount] <> null and [Discount] < 1 then [Sales Amount] / (1 - [Discount]) else [Sales Amount], type number),
    AddDiscountAmount = Table.AddColumn(AddListPrice, "Discount Amount", each [List Price] - [Sales Amount], type number),
    AddCost = Table.AddColumn(AddDiscountAmount, "Cost", each [Sales Amount] - [Profit], type number),
    AddBand = Table.AddColumn(AddCost, "Discount Band", each if [Discount] = null or [Discount] = 0 then "No discount" else if [Discount] <= 0.2 then "1% to 20%" else "Above 20%", type text),
    AddBandOrder = Table.AddColumn(AddBand, "Discount Band Order", each if [Discount Band] = "No discount" then 1 else if [Discount Band] = "1% to 20%" then 2 else 3, Int64.Type)
in
    AddBandOrder""", "Table"))

EXPRESSIONS.append(("TargetBase", "Sales targets with a month date built from Year and Month.", """let
    Sheet = Workbook{[Item = "Sales Target", Kind = "Sheet"]}[Data],
    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Headers, {{"State", type text}, {"Year", Int64.Type}, {"Month", Int64.Type}, {"Target", type number}}, "en-US"),
    Trimmed = Table.TransformColumns(Typed, {{"State", Text.Trim, type text}}),
    Valid = Table.SelectRows(Trimmed, each [State] <> null and [Year] <> null and [Month] <> null),
    AddMonthStart = Table.AddColumn(Valid, "Month Start", each #date([Year], [Month], 1), type date),
    Result = Table.SelectColumns(AddMonthStart, {"State", "Month Start", "Target"})
in
    Result""", "Table"))

# ---------------------------------------------------------------------------- tables
# column: (name, dataType, format, summarizeBy, description, options)   options: hidden, sortBy, key

TABLES = []


def T(name, doc, cols, m):
    TABLES.append({"name": name, "doc": doc, "cols": cols, "m": m})


CUR = r"\$#,0.00"
T("Orders", "One row per order line. Returns, list price, discount amount and cost are added in Power Query.", [
    ("Row ID", "int64", "0", "none", "Line number from the sheet.", {"hidden": True}),
    ("Order ID", "string", None, "none", "Order number. An order can have several lines.", {}),
    ("Order Date", "dateTime", "dd mmm yyyy", "none", "Date of the order.", {}),
    ("Ship Mode", "string", None, "none", "Shipping method.", {}),
    ("Customer ID", "string", None, "none", "Links to Customer.", {"hidden": True}),
    ("City", "string", None, "none", "City of the order.", {}),
    ("State", "string", None, "none", "Links to State.", {"hidden": True}),
    ("Product ID", "string", None, "none", "Links to Product.", {"hidden": True}),
    ("Sales Amount", "double", CUR, "sum", "Sales value of the line, after discount, before returns.", {}),
    ("Quantity", "int64", "#,0", "sum", "Units on the line.", {}),
    ("Discount", "double", "0%", "none", "Discount rate given on the line.", {}),
    ("Profit", "double", CUR, "sum", "Profit of the line.", {}),
    ("Sales Person ID", "string", None, "none", "Links to Sales Person.", {"hidden": True}),
    ("Returned", "string", None, "none", "Yes when the order is on the Returns sheet.", {}),
    ("Returned Amount", "double", CUR, "sum", "Sales Amount when the order was returned, else 0.", {}),
    ("Net Sales", "double", CUR, "sum", "Sales Amount minus Returned Amount.", {}),
    ("List Price", "double", CUR, "sum", "Value before discount: Sales Amount / (1 - Discount).", {}),
    ("Discount Amount", "double", CUR, "sum", "List Price minus Sales Amount.", {}),
    ("Cost", "double", CUR, "sum", "Sales Amount minus Profit.", {}),
    ("Discount Band", "string", None, "none", "No discount, 1% to 20%, or Above 20%.", {"sortBy": "Discount Band Order"}),
    ("Discount Band Order", "int64", "0", "none", "Sort order of Discount Band.", {"hidden": True}),
], """let
    Source = OrdersBase,
    // Segment lives in Customer and Region in State, so they are not repeated here
    Result = Table.RemoveColumns(Source, {"Segment", "Region"})
in
    Result""")

T("Sales Target", "One row per state per month. Target is a sales amount in dollars.", [
    ("State", "string", None, "none", "Links to State.", {"hidden": True}),
    ("Month Start", "dateTime", "mmm yyyy", "none", "First day of the target month.", {"hidden": True}),
    ("Target", "double", r"\$#,0", "sum", "Sales target for the state and month.", {}),
], """let
    Source = TargetBase
in
    Source""")

T("Date", "One row per day. Starts on 1 January of the first year in the data and ends on 31 December of the last year, so it grows with the data.", [
    ("Date", "dateTime", "dd mmm yyyy", "none", "Calendar day.", {"key": True}),
    ("Year", "int64", "0", "none", "Calendar year.", {}),
    ("Quarter", "string", None, "none", "Q1 to Q4.", {}),
    ("Month Number", "int64", "0", "none", "1 to 12, sorts the month names.", {"hidden": True}),
    ("Month", "string", None, "none", "Short month name.", {"sortBy": "Month Number"}),
    ("Month Start", "dateTime", "mmm yyyy", "none", "First day of the month.", {}),
    ("Year Month", "string", None, "none", "Year and month, for example 2021-03.", {}),
], """let
    AllDates = List.Combine({OrdersBase[Order Date], TargetBase[Month Start]}),
    FirstDate = List.Min(AllDates),
    LastDate = List.Max(AllDates),
    StartDate = #date(Date.Year(FirstDate), 1, 1),
    EndDate = #date(Date.Year(LastDate), 12, 31),
    Days = List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)),
    AsTable = Table.FromList(Days, Splitter.SplitByNothing(), type table [Date = date]),
    AddYear = Table.AddColumn(AsTable, "Year", each Date.Year([Date]), Int64.Type),
    AddQuarter = Table.AddColumn(AddYear, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
    AddMonthNumber = Table.AddColumn(AddQuarter, "Month Number", each Date.Month([Date]), Int64.Type),
    AddMonth = Table.AddColumn(AddMonthNumber, "Month", each Date.ToText([Date], "MMM", "en-US"), type text),
    AddMonthStart = Table.AddColumn(AddMonth, "Month Start", each Date.StartOfMonth([Date]), type date),
    AddYearMonth = Table.AddColumn(AddMonthStart, "Year Month", each Date.ToText([Date], "yyyy-MM", "en-US"), type text)
in
    AddYearMonth""")

T("Product", "One row per product. Category and Sub-Category are filled down because the sheet uses merged cells.", [
    ("Product ID", "string", None, "none", "Product code.", {"hidden": True, "key": True}),
    ("Product Name", "string", None, "none", "Product name.", {}),
    ("Category", "string", None, "none", "Furniture, Office Supplies or Technology.", {}),
    ("Sub-Category", "string", None, "none", "Product group inside the category.", {}),
], """let
    Sheet = Workbook{[Item = "Product", Kind = "Sheet"]}[Data],
    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Kept = Table.SelectColumns(Headers, {"Product ID", "Product Name", "Category", "Sub-Category"}),
    Typed = Table.TransformColumnTypes(Kept, {{"Product ID", type text}, {"Product Name", type text}, {"Category", type text}, {"Sub-Category", type text}}),
    Trimmed = Table.TransformColumns(Typed, {{"Product ID", Text.Trim, type text}, {"Product Name", Text.Trim, type text}, {"Category", Text.Trim, type text}, {"Sub-Category", Text.Trim, type text}}),
    // Merged cells arrive as blanks: turn empty text into null, then fill down
    BlankToNull = Table.ReplaceValue(Trimmed, "", null, Replacer.ReplaceValue, {"Category", "Sub-Category"}),
    FilledDown = Table.FillDown(BlankToNull, {"Category", "Sub-Category"}),
    Valid = Table.SelectRows(FilledDown, each [Product ID] <> null),
    Result = Table.Distinct(Valid, {"Product ID"})
in
    Result""")

T("Customer", "One row per customer, with the RFM scores. Recency follows the brief: sales in the 180 days before the latest order date in the data.", [
    ("Customer ID", "string", None, "none", "Customer code.", {"hidden": True, "key": True}),
    ("Customer Name", "string", None, "none", "Customer name.", {}),
    ("Segment", "string", None, "none", "Consumer, Corporate or Home Office.", {}),
    ("First Order Date", "dateTime", "dd mmm yyyy", "none", "Date of the customer's first order.", {}),
    ("First Order Year", "int64", "0", "none", "Year of the first order. A customer is new in that year.", {}),
    ("Last Order Date", "dateTime", "dd mmm yyyy", "none", "Date of the customer's latest order.", {}),
    ("Lifetime Orders", "int64", "#,0", "none", "Frequency: number of orders, all years.", {}),
    ("Lifetime Sales", "double", r"\$#,0", "none", "Monetary: sales after returns, all years.", {}),
    ("Sales Last 180 Days", "double", r"\$#,0", "none", "Recency: sales after returns in the last 180 days of the data.", {}),
    ("R Score", "int64", "0", "none", "1 to 5. Fifths of customers by Sales Last 180 Days; equal values get the same score.", {}),
    ("F Score", "int64", "0", "none", "1 to 5. Fifths of customers by Orders.", {}),
    ("M Score", "int64", "0", "none", "1 to 5. Fifths of customers by Total Sales.", {}),
    ("RFM Score", "int64", "0", "none", "R + F + M, from 3 to 15.", {}),
    ("Value Tier", "string", None, "none", "High (12 to 15), Medium (8 to 11) or Low (3 to 7).", {"sortBy": "Tier Order"}),
    ("Tier Order", "int64", "0", "none", "Sort order of Value Tier.", {"hidden": True}),
    ("Action", "string", None, "none", "Marketing action for the customer.", {"sortBy": "Action Order"}),
    ("Action Order", "int64", "0", "none", "Sort order of Action.", {"hidden": True}),
    ("Action Rule", "string", None, "none", "The rule that puts a customer in the action group.", {}),
    ("Action Advice", "string", None, "none", "What to do with the action group.", {}),
], """let
    Sheet = Workbook{[Item = "Customer", Kind = "Sheet"]}[Data],
    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Kept = Table.SelectColumns(Headers, {"Customer ID", "Customer Name"}),
    Typed = Table.TransformColumnTypes(Kept, {{"Customer ID", type text}, {"Customer Name", type text}}),
    Trimmed = Table.TransformColumns(Typed, {{"Customer ID", Text.Trim, type text}, {"Customer Name", Text.Trim, type text}}),
    Customers = Table.Distinct(Table.SelectRows(Trimmed, each [Customer ID] <> null), {"Customer ID"}),
    // The latest order date in the data is treated as "today"
    LatestDate = List.Max(OrdersBase[Order Date]),
    WindowStart = Date.AddDays(LatestDate, -180),
    Stats = Table.Group(OrdersBase, {"Customer ID"}, {
        {"Segment", each List.First([Segment]), type text},
        {"First Order Date", each List.Min([Order Date]), type date},
        {"Last Order Date", each List.Max([Order Date]), type date},
        {"Lifetime Orders", each List.Count(List.Distinct([Order ID])), Int64.Type},
        {"Lifetime Sales", each List.Sum([Net Sales]), type number},
        {"Sales Last 180 Days", each List.Sum(Table.SelectRows(_, (r) => r[Order Date] > WindowStart)[Net Sales]), type number}
    }),
    Joined = Table.NestedJoin(Customers, {"Customer ID"}, Stats, {"Customer ID"}, "Stats", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Stats", {"Segment", "First Order Date", "Last Order Date", "Lifetime Orders", "Lifetime Sales", "Sales Last 180 Days"}),
    NoNulls = Table.TransformColumns(Expanded, {{"Segment", each if _ = null then "No orders" else _, type text}, {"Lifetime Orders", each if _ = null then 0 else _, Int64.Type}, {"Lifetime Sales", each if _ = null then 0 else _, type number}, {"Sales Last 180 Days", each if _ = null then 0 else _, type number}}),
    Retyped = Table.TransformColumnTypes(NoNulls, {{"First Order Date", type date}, {"Last Order Date", type date}}),
    AddFirstYear = Table.AddColumn(Retyped, "First Order Year", each if [First Order Date] = null then null else Date.Year([First Order Date]), Int64.Type),
    Buffered = Table.Buffer(AddFirstYear),
    // Score 1 to 5: which fifth of customers the value falls in. Equal values always get the same score.
    AddScore = (t as table, valueColumn as text, scoreColumn as text) as table =>
        let
            RowCount = Table.RowCount(t),
            Sorted = List.Buffer(List.Sort(Table.Column(t, valueColumn)))
        in
            Table.AddColumn(t, scoreColumn, each Number.RoundDown(List.PositionOf(Sorted, Record.Field(_, valueColumn)) * 5 / RowCount) + 1, Int64.Type),
    WithR = AddScore(Buffered, "Sales Last 180 Days", "R Score"),
    WithF = AddScore(WithR, "Lifetime Orders", "F Score"),
    WithM = AddScore(WithF, "Lifetime Sales", "M Score"),
    AddRfm = Table.AddColumn(WithM, "RFM Score", each [R Score] + [F Score] + [M Score], Int64.Type),
    AddTier = Table.AddColumn(AddRfm, "Value Tier", each if [RFM Score] >= 12 then "High" else if [RFM Score] >= 8 then "Medium" else "Low", type text),
    AddTierOrder = Table.AddColumn(AddTier, "Tier Order", each if [Value Tier] = "High" then 1 else if [Value Tier] = "Medium" then 2 else 3, Int64.Type),
    // Bonus: one marketing action per customer
    AddActionOrder = Table.AddColumn(AddTierOrder, "Action Order", each
        if [M Score] >= 4 and [Sales Last 180 Days] = 0 then 1
        else if [Value Tier] = "High" then 2
        else if [Value Tier] = "Medium" and [Sales Last 180 Days] > 0 then 3
        else if [Sales Last 180 Days] > 0 then 4
        else 5, Int64.Type),
    AddAction = Table.AddColumn(AddActionOrder, "Action", each {"Win back", "Reward and keep", "Grow", "Nurture", "Lapsed, low value"}{[Action Order] - 1}, type text),
    AddRule = Table.AddColumn(AddAction, "Action Rule", each {"M score 4 or 5, nothing bought in the last 180 days", "High value tier", "Medium tier, bought in the last 180 days", "Low tier, bought in the last 180 days", "Everyone else with nothing bought in the last 180 days"}{[Action Order] - 1}, type text),
    AddAdvice = Table.AddColumn(AddRule, "Action Advice", each {"Personal call or offer from the sales person", "Loyalty benefits, no deep discounts", "Cross-sell categories they have not bought", "Low-cost email campaigns", "No spend for now"}{[Action Order] - 1}, type text)
in
    AddAdvice""")

T("State", "One row per state, with its region and the region head.", [
    ("State", "string", None, "none", "State name.", {"key": True}),
    ("Region", "string", None, "none", "Central, East, South or West.", {}),
    ("Region Head", "string", None, "none", "Person responsible for the region.", {}),
], """let
    FromOrders = Table.Distinct(Table.SelectColumns(OrdersBase, {"State", "Region"}), {"State"}),
    // A state that has a target but no orders yet still needs a row
    KnownStates = List.Buffer(FromOrders[State]),
    TargetOnly = Table.SelectRows(Table.Distinct(Table.SelectColumns(TargetBase, {"State"})), each not List.Contains(KnownStates, [State])),
    TargetOnlyWithRegion = Table.AddColumn(TargetOnly, "Region", each "Not assigned", type text),
    AllStates = Table.Combine({FromOrders, TargetOnlyWithRegion}),
    HeadsSheet = Table.PromoteHeaders(Workbook{[Item = "Region Heads", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
    Heads = Table.TransformColumns(Table.RenameColumns(Table.SelectColumns(HeadsSheet, {"Person", "Region"}), {{"Person", "Region Head"}}), {{"Region Head", each Text.Trim(Text.From(_)), type text}, {"Region", each Text.Trim(Text.From(_)), type text}}),
    Joined = Table.NestedJoin(AllStates, {"Region"}, Heads, {"Region"}, "Head", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Head", {"Region Head"}),
    NoNulls = Table.TransformColumns(Expanded, {{"Region Head", each if _ = null then "Not assigned" else _, type text}}),
    Result = Table.Sort(NoNulls, {{"State", Order.Ascending}})
in
    Result""")

T("Sales Person", "One row per sales person.", [
    ("Sales Person ID", "string", None, "none", "Sales person code.", {"hidden": True, "key": True}),
    ("Sales Person", "string", None, "none", "Sales person name.", {}),
], """let
    Sheet = Workbook{[Item = "SalesPerson", Kind = "Sheet"]}[Data],
    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Renamed = Table.RenameColumns(Headers, {{"Sales_person_ID", "Sales Person ID"}, {"Sales_person", "Sales Person"}}),
    Typed = Table.TransformColumnTypes(Renamed, {{"Sales Person ID", type text}, {"Sales Person", type text}}),
    Trimmed = Table.TransformColumns(Typed, {{"Sales Person ID", Text.Trim, type text}, {"Sales Person", Text.Trim, type text}}),
    Result = Table.Distinct(Table.SelectRows(Trimmed, each [Sales Person ID] <> null), {"Sales Person ID"})
in
    Result""")

T("Trend Year", "The last 4 years in the data, for trend charts. Not linked to other tables, so the Year slicer does not hide the earlier years.", [
    ("Year", "int64", "0", "none", "Calendar year.", {}),
], """let
    Years = List.Sort(List.Distinct(List.Transform(OrdersBase[Order Date], Date.Year))),
    LastFour = List.LastN(Years, 4),
    Result = Table.FromList(LastFour, Splitter.SplitByNothing(), type table [Year = Int64.Type])
in
    Result""")


def series_table(name, col, doc, m_list, mtype="Int64.Type", dtype="int64", fmt="0"):
    T(name, doc, [(col, dtype, fmt, "none", doc, {})], f"""let
    Values = {m_list},
    Result = Table.FromList(Values, Splitter.SplitByNothing(), type table [#"{col}" = {mtype}])
in
    Result""")


for seg in ("Consumer", "Corporate", "Home Office"):
    series_table(f"Discount {seg}", f"{seg} Discount %", f"What-if input: discount % to test for the {seg} segment, 0 to 50.", "List.Numbers(0, 51, 1)")
series_table("Volume Response", "Extra Units %", "What-if input: % more units sold for every 1 point of extra discount (and fewer units for every point removed).",
             "List.Numbers(0, 11, 0.5)", "number", "double", "0.0")
series_table("Top Customer %", "Top %", "Pareto input: the top % of customers to measure, 1 to 100.", "List.Numbers(1, 100, 1)")
series_table("Pareto Axis", "Top % of Customers", "Axis of the Pareto curve, 1 to 100.", "List.Numbers(1, 100, 1)")

T("Top N", "Choices for how many most-returned products to show: 5, 10, 15 and so on, up to the number of returned products in the data.", [
    ("Top N", "int64", "0", "none", "Number of products to show.", {}),
], """let
    ReturnedProducts = List.Count(List.Distinct(Table.SelectRows(OrdersBase, each [Returned] = "Yes")[Product ID])),
    Steps = List.Max({Number.RoundUp(ReturnedProducts / 5), 1}),
    Values = List.Numbers(5, Steps, 5),
    Result = Table.FromList(Values, Splitter.SplitByNothing(), type table [#"Top N" = Int64.Type])
in
    Result""")

T("Sales Basis", "Choice for the target page: compare the target with sales after returns or before returns.", [
    ("Basis", "string", None, "none", "After returns or Before returns.", {"sortBy": "Basis Order"}),
    ("Basis Order", "int64", "0", "none", "Sort order of Basis.", {"hidden": True}),
], """let
    Result = #table(type table [Basis = text, #"Basis Order" = Int64.Type], {{"After returns", 1}, {"Before returns", 2}})
in
    Result""")

T("Insight", "The five headline insights listed on the first page. The sentences are measures, so they follow the slicers.", [
    ("No", "int64", "0", "none", "Insight number.", {}),
    ("Theme", "string", None, "none", "Short name of the insight.", {"sortBy": "No"}),
], """let
    Result = #table(type table [No = Int64.Type, Theme = text], {{1, "Growth and target"}, {2, "Returns"}, {3, "Targets by state"}, {4, "Discounts"}, {5, "Customers"}})
in
    Result""")

RELATIONSHIPS = [
    ("Orders", "Order Date", "Date", "Date"),
    ("Orders", "Customer ID", "Customer", "Customer ID"),
    ("Orders", "Product ID", "Product", "Product ID"),
    ("Orders", "State", "State", "State"),
    ("Orders", "Sales Person ID", "Sales Person", "Sales Person ID"),
    ("Sales Target", "State", "State", "State"),
    ("Sales Target", "Month Start", "Date", "Date"),
]

# ---------------------------------------------------------------------------- measures

MEASURES = []  # name, dax, format, folder, description


def M(name, dax, fmt, folder, desc):
    MEASURES.append((name, dax.strip("\n"), fmt, folder, desc))


def cy(expr):
    return f"VAR y = [Selected Year]\nRETURN\n    CALCULATE ( {expr}, 'Date'[Year] = y )"


def k(expr):
    """DAX text: a dollar amount in thousands, for example $657.7K."""
    return f'FORMAT ( DIVIDE ( {expr}, 1000 ), "\\$#,0.0" ) & "K"'


USD, USD2, PCT, INT = r"\$#,0", r"\$#,0.00", "0.0%", "#,0"

M("Latest Year", "YEAR ( CALCULATE ( MAX ( Orders[Order Date] ), REMOVEFILTERS () ) )", "0", "Base",
  "Last year that has orders in the data.")
M("Selected Year", "COALESCE ( SELECTEDVALUE ( 'Date'[Year] ), [Latest Year] )", "0", "Base",
  "The year picked in the Year slicer. With nothing picked it is the latest year in the data, so the report moves on by itself when a new year arrives.")
M("Sales Before Returns", cy("SUM ( Orders[Sales Amount] )"), USD, "Base", "Sales Amount of all order lines in the selected year.")
M("Returned Sales", cy("SUM ( Orders[Returned Amount] )"), USD, "Base", "Sales Amount of order lines whose order was returned.")
M("Total Sales", cy("SUM ( Orders[Net Sales] )"), USD, "Base", "Required KPI: Sales minus Returns, selected year.")
M("Profit", cy("SUM ( Orders[Profit] )"), USD, "Base", "Profit of all order lines in the selected year.")
M("% Margin", "DIVIDE ( [Profit], [Sales Before Returns] )", PCT, "Base", "Required KPI: Profit divided by Sales Amount.")
M("Quantity", cy("SUM ( Orders[Quantity] )"), INT, "Base", "Units sold.")
M("Orders", cy("DISTINCTCOUNT ( Orders[Order ID] )"), INT, "Base", "Number of orders.")
M("Order Lines", cy("COUNTROWS ( Orders )"), INT, "Base", "Number of order lines.")
M("Customers", cy("DISTINCTCOUNT ( Orders[Customer ID] )"), INT, "Base", "Customers with at least one order in the selected year.")
M("Discount Amount", cy("SUM ( Orders[Discount Amount] )"), USD, "Base", "Value given away as discount: list price minus sales amount.")
M("List Sales", cy("SUM ( Orders[List Price] )"), USD, "Base", "Sales value before any discount.")
M("% Avg Discount", cy("AVERAGE ( Orders[Discount] )"), PCT, "Base", "Simple average of the discount rate over order lines.")
M("% Discount Given", "DIVIDE ( [Discount Amount], [List Sales] )", PCT, "Base", "Discount amount as a share of list price (weighted by value).")

for m, fmt in (("Total Sales", USD), ("Profit", USD), ("Orders", INT), ("Customers", INT), ("Sales Before Returns", USD), ("Returned Sales", USD)):
    M(f"{m} LY", f"VAR y = [Selected Year]\nRETURN\n    CALCULATE ( [{m}], 'Date'[Year] = y - 1 )", fmt, "Last year", f"{m} in the year before the selected year.")
for m in ("Total Sales", "Profit", "Orders", "Customers"):
    M(f"{m} YoY %", f"VAR ly = [{m} LY]\nRETURN\n    IF ( ly <> 0, DIVIDE ( [{m}] - ly, ly ) )", "+0.0%;-0.0%;0.0%", "Last year", f"Growth of {m} on the year before.")
    M(f"{m} YoY Text",
      f'VAR g = [{m} YoY %]\nRETURN\n    IF ( ISBLANK ( g ), "No data for " & ( [Selected Year] - 1 ), IF ( g >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( g ), "0.0%" ) & " vs " & ( [Selected Year] - 1 ) )',
      None, "Card text", f"Small line under the {m} card.")
M("% Margin LY", "VAR y = [Selected Year]\nRETURN\n    CALCULATE ( [% Margin], 'Date'[Year] = y - 1 )", PCT, "Last year", "% Margin in the year before.")
M("% Margin Text", 'VAR ly = [% Margin LY]\nRETURN\n    IF ( ISBLANK ( ly ), "No data for " & ( [Selected Year] - 1 ), FORMAT ( ly, "0.0%" ) & " in " & ( [Selected Year] - 1 ) )',
  None, "Card text", "Small line under the % Margin card.")
M("Year Shown", '"Year shown: " & [Selected Year]', None, "Card text", "Tells the reader which year the page is showing.")

# trend (last 4 years)
M("Sales Trend", "VAR ty = SELECTEDVALUE ( 'Trend Year'[Year] )\nRETURN\n    IF ( NOT ISBLANK ( ty ), CALCULATE ( SUM ( Orders[Net Sales] ), 'Date'[Year] = ty ) )", USD, "Trend",
  "Total sales for a year of the Trend Year table. Used on the 4-year charts, which ignore the Year slicer.")
M("Return Rate Trend", "VAR ty = SELECTEDVALUE ( 'Trend Year'[Year] )\nRETURN\n    IF (\n        NOT ISBLANK ( ty ),\n        DIVIDE (\n            CALCULATE ( SUM ( Orders[Returned Amount] ), 'Date'[Year] = ty ),\n            CALCULATE ( SUM ( Orders[Sales Amount] ), 'Date'[Year] = ty )\n        )\n    )",
  PCT, "Trend", "Return rate for a year of the Trend Year table.")

# target
M("Target", """VAR y = [Selected Year]
RETURN
    IF (
        NOT ( ISFILTERED ( Customer ) || ISFILTERED ( Product ) || ISFILTERED ( 'Sales Person' ) || ISFILTERED ( Orders[City] ) ),
        CALCULATE ( SUM ( 'Sales Target'[Target] ), 'Date'[Year] = y )
    )""", USD, "Target", "Sales target for the selected year. Targets exist per state and month only, so it is blank when a segment, category, sales person or city is selected.")
M("Sales For Target", 'IF ( SELECTEDVALUE ( \'Sales Basis\'[Basis], "After returns" ) = "After returns", [Total Sales], [Sales Before Returns] )', USD, "Target",
  "Sales compared with the target: after returns unless the Sales basis slicer says before returns.")
M("% Target Achieved", "DIVIDE ( [Sales For Target] + 0, [Target] )", PCT, "Target", "Required KPI: sales divided by target.")
M("Target Gap", "IF ( NOT ISBLANK ( [Target] ), [Sales For Target] - [Target] )", r"\$#,0;-\$#,0;\$0", "Target", "Sales minus target. Negative means the target was missed.")
M("Target Status", 'IF ( NOT ISBLANK ( [Target] ), IF ( [% Target Achieved] >= 1, "Achieved", "Not achieved" ) )', None, "Target", "Achieved when sales reach the target.")
M("Target Colour", 'IF ( ISBLANK ( [Target] ), "#111827", IF ( [% Target Achieved] >= 1, "#15803D", "#C2410C" ) )', None, "Target", "Green when the target is achieved, orange-red when not.")
M("Target Text", 'VAR g = [Target Gap]\nRETURN\n    IF ( ISBLANK ( [Target] ), "Targets are set by state only", IF ( g >= 0, "Above target by ", "Short of target by " ) & ' + k("ABS ( g )") + ' )', None, "Card text", "Small line under the % Target Achieved card.")
M("Target Card Text", '"Target " & ' + k("[Target]"), None, "Card text", "Small line showing the target.")
M("States Achieved", "COUNTROWS ( FILTER ( VALUES ( State[State] ), NOT ISBLANK ( [Target] ) && [% Target Achieved] >= 1 ) ) + 0", INT, "Target", "States that reached their target.")
M("States With Target", "COUNTROWS ( FILTER ( VALUES ( State[State] ), NOT ISBLANK ( [Target] ) ) ) + 0", INT, "Target", "States that have a target.")
M("States Achieved Text", '[States Achieved] & " of " & [States With Target]', None, "Target", "For example 34 of 49.")
M("States Missed Text", '( [States With Target] - [States Achieved] ) & " states missed target"', None, "Card text", "Small line under the states card.")
M("Regions Achieved", "SUMX ( VALUES ( State[Region] ), IF ( NOT ISBLANK ( [Target] ) && [% Target Achieved] >= 1, 1, 0 ) )", INT, "Target", "Regions that reached their target.")
M("Regions With Target", "SUMX ( VALUES ( State[Region] ), IF ( NOT ISBLANK ( [Target] ), 1, 0 ) )", INT, "Target", "Regions that have a target.")
M("Regions Achieved Text", '[Regions Achieved] & " of " & [Regions With Target]', None, "Target", "For example 3 of 4.")
M("Months Achieved Text",
  'VAR t = FILTER ( VALUES ( \'Date\'[Month Number] ), NOT ISBLANK ( [Target] ) )\nRETURN\n    COUNTROWS ( FILTER ( t, [% Target Achieved] >= 1 ) ) + 0 & " of " & COUNTROWS ( t ) + 0 & " months on target"',
  None, "Card text", "Small line: months that reached the monthly target.")
M("Basis Text", '"Sales " & LOWER ( SELECTEDVALUE ( \'Sales Basis\'[Basis], "After returns" ) )', None, "Card text", "Small line: which sales the target is compared with.")

# returns
M("Return Rate %", "DIVIDE ( [Returned Sales], [Sales Before Returns] )", PCT, "Returns", "Returned sales as a share of sales before returns.")
M("Return Rate LY", "VAR y = [Selected Year]\nRETURN\n    CALCULATE ( [Return Rate %], 'Date'[Year] = y - 1 )", PCT, "Returns", "Return rate in the year before.")
M("Return Rate Text", 'VAR ly = [Return Rate LY]\nRETURN\n    IF ( ISBLANK ( ly ), "No data for " & ( [Selected Year] - 1 ), FORMAT ( ly, "0.0%" ) & " in " & ( [Selected Year] - 1 ) )', None, "Card text", "Small line under the return rate card.")
M("Returned Sales Text", 'VAR ly = [Returned Sales LY]\nRETURN\n    IF ( ISBLANK ( ly ), "No data for " & ( [Selected Year] - 1 ), ' + k("ly") + ' & " in " & ( [Selected Year] - 1 ) )', None, "Card text", "Small line under the returned sales card.")
M("Returned Orders", cy("DISTINCTCOUNT ( Orders[Order ID] ), Orders[Returned] = \"Yes\"").replace("), Orders", "), Orders"), INT, "Returns", "Orders that were returned.")
M("Returned Orders Text", '"of " & FORMAT ( [Orders], "#,0" ) & " orders"', None, "Card text", "Small line under the returned orders card.")
M("Returned Quantity", cy("SUM ( Orders[Quantity] ), Orders[Returned] = \"Yes\""), INT, "Returns", "Units on returned orders.")
M("Returned Quantity Text", '"of " & FORMAT ( [Quantity], "#,0" ) & " units sold"', None, "Card text", "Small line under the returned quantity card.")
M("Returned Profit", cy("SUM ( Orders[Profit] ), Orders[Returned] = \"Yes\""), USD, "Returns", "Profit that was booked on returned orders.")
M("Top N Value", "SELECTEDVALUE ( 'Top N'[Top N], 5 )", "0", "Returns", "Number of products picked in the Top N slicer (5 when nothing is picked).")
M("Return Rank", "IF ( [Returned Sales] > 0, RANKX ( ALLSELECTED ( Product[Product Name] ), [Returned Sales],, DESC, SKIP ) )", "0", "Returns", "Rank of a product by returned sales, 1 = most returned.")
for name, expr, fmt, desc in (
        ("Returned Sales Top N", "[Returned Sales]", USD, "Returned sales, only for products inside the chosen Top N."),
        ("Returned Quantity Top N", "[Returned Quantity]", INT, "Returned units, only for products inside the chosen Top N."),
        ("Returned Orders Top N", "[Returned Orders]", INT, "Returned orders, only for products inside the chosen Top N."),
        ("Return Rank Top N", "[Return Rank]", "0", "Rank, only for products inside the chosen Top N."),
        ("Returned Category", 'CONCATENATEX ( VALUES ( Product[Category] ), Product[Category], ", " )', None, "Category of the returned product."),
        ("Returned Sub-Category", 'CONCATENATEX ( VALUES ( Product[Sub-Category] ), Product[Sub-Category], ", " )', None, "Sub-category of the returned product."),
        ("Top Return Region",
         'CONCATENATEX ( TOPN ( 1, ADDCOLUMNS ( VALUES ( State[Region] ), "@r", [Returned Sales] ), [@r], DESC ), State[Region], ", " )', None,
         "Region where most of the product's returns happened.")):
    M(name, f"IF ( NOT ISBLANK ( [Return Rank] ) && [Return Rank] <= [Top N Value], {expr} )", fmt, "Returns", desc)
M("Top N Title", '"Top " & [Top N Value] & " most-returned products, " & [Selected Year]', None, "Returns", "Chart title that follows the Top N slicer.")

# what-if
WHATIF = """VAR up = SELECTEDVALUE ( 'Volume Response'[Extra Units %], 0 ) / 100
RETURN
    SUMX (
        VALUES ( Customer[Segment] ),
        VAR seg = Customer[Segment]
        VAR listValue = [List Sales]
        VAR sales = [Sales Before Returns]
        VAR cost = sales - [Profit]
        VAR currentDiscount = DIVIDE ( listValue - sales, listValue )
        VAR picked =
            SWITCH (
                seg,
                "Consumer", SELECTEDVALUE ( 'Discount Consumer'[Consumer Discount %] ),
                "Corporate", SELECTEDVALUE ( 'Discount Corporate'[Corporate Discount %] ),
                "Home Office", SELECTEDVALUE ( 'Discount Home Office'[Home Office Discount %] )
            )
        VAR newDiscount = IF ( ISBLANK ( picked ), currentDiscount, picked / 100 )
        VAR volume = 1 + up * ( newDiscount - currentDiscount ) * 100
        RETURN
            __RESULT__
    )"""
M("Predicted Sales", WHATIF.replace("__RESULT__", "listValue * ( 1 - newDiscount ) * volume"), USD, "What-if",
  "Sales if each segment's discount is set to the slider value: list price x (1 - new discount), with the optional volume response.")
M("Predicted Profit", WHATIF.replace("__RESULT__", "( listValue * ( 1 - newDiscount ) - cost ) * volume"), USD, "What-if",
  "Profit at the new discounts: predicted sales minus cost.")
M("Tested Discount %", "DIVIDE (\n    " + WHATIF.replace("__RESULT__", "listValue * newDiscount").replace("\n", "\n    ") + ",\n    [List Sales]\n)", PCT, "What-if",
  "The discount being tested, weighted by list price.")
M("Predicted Profit %", "DIVIDE ( [Predicted Profit], [Predicted Sales] )", PCT, "What-if", "Required output of the what-if: predicted profit divided by predicted sales.")
M("Profit % Change", "( [Predicted Profit %] - [% Margin] ) * 100", '+0.0;-0.0;0.0', "What-if", "Predicted profit % minus current profit %, in percentage points.")
M("Predicted Profit Change", "[Predicted Profit] - [Profit]", r"+\$#,0;-\$#,0;\$0", "What-if", "Predicted profit minus current profit.")
M("Predicted Profit Text", 'VAR d = [Predicted Profit Change]\nRETURN\n    IF ( d >= 0, "▲ ", "▼ " ) & ' + k("ABS ( d )") + ' & " vs current " & ' + k("[Profit]"), None, "Card text", "Small line under the predicted profit card.")
M("Predicted Sales Text", 'VAR d = [Predicted Sales] - [Sales Before Returns]\nRETURN\n    IF ( d >= 0, "▲ ", "▼ " ) & ' + k("ABS ( d )") + ' & " vs current " & ' + k("[Sales Before Returns]"), None, "Card text", "Small line under the predicted sales card.")
M("Current Profit Text", k("[Profit]") + ' & " profit on " & ' + k("[Sales Before Returns]") + ' & " sales"', None, "Card text", "Small line under the current profit % card.")
M("Tested Discount Text", '"Current discount " & FORMAT ( [% Discount Given], "0.0%" )', None, "Card text", "Small line under the tested discount card.")
M("Profit Colour", 'IF ( [Profit] >= 0, "#1D4ED8", "#C2410C" )', None, "What-if", "Blue for a profit, orange-red for a loss.")

# pareto
PARETO = """VAR n = ROUNDUP ( [Customers] * __PCT__ / 100, 0 )
VAR t = ADDCOLUMNS ( VALUES ( Customer[Customer ID] ), "@s", [Total Sales] )
RETURN
    SUMX ( TOPN ( n, t, [@s], DESC ), [@s] )"""
M("Top Pct", "SELECTEDVALUE ( 'Top Customer %'[Top %], 20 )", "0", "Pareto", "The top % of customers picked by the user (20 when nothing is picked).")
M("Top Customers Count", "ROUNDUP ( [Customers] * [Top Pct] / 100, 0 )", INT, "Pareto", "How many customers are in the chosen top %.")
M("Top Customers Sales", "VAR p = [Top Pct]\n" + PARETO.replace("__PCT__", "p"), USD, "Pareto", "Total sales of the customers in the chosen top %.")
M("Top Customers Share %", "DIVIDE ( [Top Customers Sales], [Total Sales] )", PCT, "Pareto", "Required output of the Pareto view: share of total sales from the chosen top % of customers.")
M("Top Customers Count Text", 'FORMAT ( [Top Customers Count], "#,0" ) & " of " & FORMAT ( [Customers], "#,0" ) & " customers"', None, "Card text", "Small line: customers in the chosen top %.")
M("Top Customers Sales Text", '"of " & ' + k("[Total Sales]") + ' & " total sales"', None, "Card text", "Small line under the top customers sales card.")
M("Top Share Title", '"Sales from the top " & [Top Pct] & "% of customers"', None, "Pareto", "Card title that follows the slider.")
M("Pareto Curve %", "VAR p = SELECTEDVALUE ( 'Pareto Axis'[Top % of Customers] )\nRETURN\n    IF (\n        NOT ISBLANK ( p ),\n        DIVIDE (\n            " + PARETO.replace("__PCT__", "p").replace("\n", "\n            ") + ",\n            [Total Sales]\n        )\n    )",
  PCT, "Pareto", "Share of sales from the top x% of customers, for each x on the Pareto Axis.")
M("Pareto Selected Point", "IF ( SELECTEDVALUE ( 'Pareto Axis'[Top % of Customers] ) = [Top Pct], [Pareto Curve %] )", PCT, "Pareto", "The point of the curve at the chosen top %, drawn as a marker.")
M("Customers For 80%", """VAR t = ADDCOLUMNS ( VALUES ( Customer[Customer ID] ), "@s", [Total Sales] )
VAR total = SUMX ( t, [@s] )
VAR withBefore = ADDCOLUMNS ( t, "@before", VAR s = [@s] RETURN SUMX ( FILTER ( t, [@s] > s ), [@s] ) )
RETURN
    COUNTROWS ( FILTER ( withBefore, [@s] > 0 && [@before] < 0.8 * total ) ) + 0""", INT, "Pareto", "Smallest number of customers, biggest first, that add up to 80% of sales.")
M("Customers For 80% Text", 'FORMAT ( DIVIDE ( [Customers For 80%], [Customers] ), "0.0%" ) & " of customers"', None, "Card text", "Small line under the 80% card.")
M("Customer Rank", "VAR cur = [Total Sales]\nRETURN\n    IF ( NOT ISBLANK ( cur ), COUNTROWS ( FILTER ( ALLSELECTED ( Customer[Customer ID] ), [Total Sales] > cur ) ) + 1 )", "0", "Pareto", "Rank of a customer by total sales in the selected year.")
M("Top Customers Count All", "CALCULATE ( [Top Customers Count], ALLSELECTED ( Customer ) )", INT, "Pareto", "Top Customers Count for all customers in view, usable on a customer row.")
M("Top Customer Rank", "IF ( [Customer Rank] <= [Top Customers Count All], [Customer Rank] )", "0", "Pareto", "Rank, only for customers inside the chosen top %.")
M("Top Customer Sales", "IF ( [Customer Rank] <= [Top Customers Count All], [Total Sales] )", USD, "Pareto", "Sales, only for customers inside the chosen top %.")
M("Top Customer Share", "IF ( [Customer Rank] <= [Top Customers Count All], DIVIDE ( [Total Sales], CALCULATE ( [Total Sales], ALLSELECTED ( Customer ) ) ) )", "0.00%", "Pareto", "Share of total sales, only for customers inside the chosen top %.")
M("Top Customer Orders", "IF ( [Customer Rank] <= [Top Customers Count All], [Orders] )", INT, "Pareto", "Orders, only for customers inside the chosen top %.")

# customers: RFM (all years)
M("RFM Customers", "COUNTROWS ( Customer )", INT, "RFM", "Number of customers.")
M("RFM Sales", "SUM ( Customer[Lifetime Sales] )", USD, "RFM", "Sales after returns, all years.")
M("RFM Orders", "SUM ( Customer[Lifetime Orders] )", INT, "RFM", "Orders, all years.")
M("RFM Avg Orders", "AVERAGE ( Customer[Lifetime Orders] )", "0.0", "RFM", "Average orders per customer, all years.")
M("RFM Sales Last 180 Days", "SUM ( Customer[Sales Last 180 Days] )", USD, "RFM", "Sales in the last 180 days of the data.")
M("RFM % of Customers", "DIVIDE ( [RFM Customers], CALCULATE ( [RFM Customers], ALLSELECTED ( Customer ) ) )", PCT, "RFM", "Share of the customers in view.")
M("RFM % of Sales", "DIVIDE ( [RFM Sales], CALCULATE ( [RFM Sales], ALLSELECTED ( Customer ) ) )", PCT, "RFM", "Required: contribution to sales, as a share of the sales in view.")
M("High Value Share %", 'DIVIDE ( CALCULATE ( [RFM Sales], KEEPFILTERS ( Customer[Value Tier] = "High" ) ), [RFM Sales] )', PCT, "RFM", "Share of sales that comes from High value customers.")
M("High Value Text", 'FORMAT ( CALCULATE ( [RFM Customers], KEEPFILTERS ( Customer[Value Tier] = "High" ) ) + 0, "#,0" ) & " high value customers"', None, "Card text", "Small line under the high value card.")
M("Win Back Customers", 'CALCULATE ( [RFM Customers], KEEPFILTERS ( Customer[Action] = "Win back" ) ) + 0', INT, "RFM", "Big spenders who bought nothing in the last 180 days.")
M("Win Back Sales", 'CALCULATE ( [RFM Sales], KEEPFILTERS ( Customer[Action] = "Win back" ) )', USD, "RFM", "All-years sales of the win-back customers.")
M("Win Back Text", 'FORMAT ( DIVIDE ( [Win Back Sales], [RFM Sales] ), "0.0%" ) & " of sales, nothing in 180 days"', None, "Card text", "Small line under the win-back card.")
M("RFM Customers Text", '"Scored on " & FORMAT ( CALCULATE ( MAX ( Orders[Order Date] ), REMOVEFILTERS () ), "d mmm yyyy" )', None, "Card text", "Small line: the date the scores are measured on.")
M("RFM Sales Text", '"All years, after returns"', None, "Card text", "Small line under the sales card.")
M("RFM 180 Text", 'FORMAT ( DIVIDE ( CALCULATE ( [RFM Customers], KEEPFILTERS ( Customer[Sales Last 180 Days] > 0 ) ), [RFM Customers] ), "0%" ) & " of customers bought"', None, "Card text", "Small line: share of customers who bought in the last 180 days.")
M("Heat Colour", """VAR share = DIVIDE ( [RFM Customers], CALCULATE ( [RFM Customers], ALLSELECTED ( Customer ) ) )
RETURN
    SWITCH ( TRUE (), ISBLANK ( share ), "#FFFFFF", share >= 0.08, "#1D4ED8", share >= 0.05, "#5B84E8", share >= 0.03, "#A9BFF5", "#DCE6FB" )""", None, "RFM", "Cell colour of the score heat map: darker means more customers.")
M("Heat Font", "VAR share = DIVIDE ( [RFM Customers], CALCULATE ( [RFM Customers], ALLSELECTED ( Customer ) ) )\nRETURN\n    IF ( share >= 0.05, \"#FFFFFF\", \"#111827\" )", None, "RFM", "Text colour of the score heat map.")

# new customers
M("New Customers", "VAR y = [Selected Year]\nRETURN\n    CALCULATE ( DISTINCTCOUNT ( Orders[Customer ID] ), Customer[First Order Year] = y, 'Date'[Year] = y ) + 0", INT, "New customers",
  "Customers whose first ever order is in the selected year.")
for name, base, fmt, desc in (("New Customer Sales", "Total Sales", USD, "Total sales from customers who are new in the selected year."),
                              ("New Customer Quantity", "Quantity", INT, "Units bought by new customers."),
                              ("New Customer Discount Amount", "Discount Amount", USD, "Discount amount given to new customers."),
                              ("New Customer Orders", "Orders", INT, "Orders placed by new customers.")):
    M(name, f"VAR y = [Selected Year]\nRETURN\n    CALCULATE ( [{base}], KEEPFILTERS ( Customer[First Order Year] = y ) )", fmt, "New customers", desc)
M("New Customers Text", '"of " & FORMAT ( [Customers], "#,0" ) & " customers in " & [Selected Year]', None, "Card text", "Small line under the new customers card.")
M("New Customer Sales Text", 'FORMAT ( DIVIDE ( [New Customer Sales], [Total Sales] ), "0.0%" ) & " of total sales"', None, "Card text", "Small line under the new customer sales card.")
M("New Customer Quantity Text", '"in " & FORMAT ( [New Customer Orders], "#,0" ) & " orders"', None, "Card text", "Small line under the quantity card.")
M("New Customer Discount Text", 'VAR y = [Selected Year]\nRETURN\n    FORMAT ( DIVIDE ( [New Customer Discount Amount], CALCULATE ( [List Sales], KEEPFILTERS ( Customer[First Order Year] = y ) ) ), "0.0%" ) & " of list price"', None, "Card text", "Small line under the discount amount card.")
M("Retention %", """VAR y = [Selected Year]
VAR previousCustomers = CALCULATETABLE ( VALUES ( Orders[Customer ID] ), 'Date'[Year] = y - 1 )
VAR currentCustomers = CALCULATETABLE ( VALUES ( Orders[Customer ID] ), 'Date'[Year] = y )
RETURN
    DIVIDE ( COUNTROWS ( INTERSECT ( previousCustomers, currentCustomers ) ), COUNTROWS ( previousCustomers ) )""", PCT, "New customers", "Share of last year's customers who bought again in the selected year.")
M("Retention Text", '"of " & FORMAT ( [Customers LY], "#,0" ) & " customers from " & ( [Selected Year] - 1 )', None, "Card text", "Small line under the retention card.")
M("Sales Per Customer", "DIVIDE ( [Total Sales], [Customers] )", USD, "New customers", "Average total sales per customer.")
M("Sales Per Customer Text", '"New customers: " & FORMAT ( DIVIDE ( [New Customer Sales], [New Customers] ), "\\$#,0" )', None, "Card text", "Small line: average sales of a new customer.")

# sales people
M("Sales Person Rank", "IF ( NOT ISBLANK ( [Total Sales] ), RANKX ( ALLSELECTED ( 'Sales Person'[Sales Person] ), [Total Sales],, DESC, SKIP ) )", "0", "Sales team",
  "Rank of a sales person by total sales. With Segment on the rows it ranks inside each segment.")
for name, expr, fmt, desc in (("Top 5 Rank", "[Sales Person Rank]", "0", "Rank, only for the top 5 sales people."),
                              ("Top 5 Sales", "[Total Sales]", USD, "Total sales, only for the top 5 sales people."),
                              ("Top 5 Avg Discount", "[% Avg Discount]", PCT, "% average discount, only for the top 5 sales people."),
                              ("Top 5 Order Lines", "[Order Lines]", INT, "Order lines, only for the top 5 sales people."),
                              ("Top 5 Margin", "[% Margin]", PCT, "% margin, only for the top 5 sales people.")):
    M(name, f"IF ( [Sales Person Rank] <= 5, {expr} )", fmt, "Sales team", desc)

# insights (first page)
M("Insight 1", """VAR y = [Selected Year]
VAR g = [Total Sales YoY %]
VAR gap = [Target Gap]
RETURN
    "Total sales (after returns) were " & """ + k("[Total Sales]") + """ & " in " & y
        & IF ( ISBLANK ( g ), "", ", " & IF ( g >= 0, "up ", "down " ) & FORMAT ( ABS ( g ), "0.0%" ) & " on " & ( y - 1 ) )
        & IF (
            ISBLANK ( [Target] ),
            ".",
            ". That is " & FORMAT ( [% Target Achieved], "0.0%" ) & " of the " & """ + k("[Target]") + """ & " target: "
                & """ + k("ABS ( gap )") + """ & IF ( gap >= 0, " above it.", " short of it." )
        )""", None, "Insights", "Insight 1: growth and target.")
M("Insight 2", """VAR y = [Selected Year]
VAR worst = TOPN ( 1, ADDCOLUMNS ( VALUES ( State[Region] ), "@r", [Return Rate %] ), [@r], DESC )
VAR ly = [Return Rate LY]
RETURN
    "Returns removed " & """ + k("[Returned Sales]") + """ & ", or " & FORMAT ( [Return Rate %], "0.0%" ) & " of sales"
        & IF ( ISBLANK ( ly ), "", " (" & FORMAT ( ly, "0.0%" ) & " in " & ( y - 1 ) & ")" )
        & ". " & CONCATENATEX ( worst, State[Region], ", " ) & " has the highest return rate at " & FORMAT ( MAXX ( worst, [@r] ), "0.0%" )
        & "."
        & IF ( NOT ISBLANK ( [Target] ), " Before returns, sales were " & FORMAT ( DIVIDE ( [Sales Before Returns], [Target] ), "0.0%" ) & " of target." )""", None, "Insights", "Insight 2: returns.")
M("Insight 3", """VAR states = ADDCOLUMNS ( FILTER ( VALUES ( State[State] ), NOT ISBLANK ( [Target] ) ), "@gap", [Target Gap], "@ach", [% Target Achieved] )
VAR worst = TOPN ( 1, states, [@gap], ASC )
RETURN
    IF (
        ISBLANK ( [Target] ),
        "Targets are set per state and month, so this point is not shown while a segment or category is selected.",
        [States Achieved Text] & " states and " & [Regions Achieved Text] & " regions reached their target. The biggest gap is "
        & CONCATENATEX ( worst, State[State], ", " ) & ": " & """ + k("ABS ( MINX ( worst, [@gap] ) )") + """ & " short, at "
        & FORMAT ( MINX ( worst, [@ach] ), "0.0%" ) & " of target."
    )""", None, "Insights", "Insight 3: targets by state.")
M("Insight 4", """VAR deepProfit = CALCULATE ( [Profit], KEEPFILTERS ( Orders[Discount Band] = "Above 20%" ) )
VAR deepSales = CALCULATE ( [Sales Before Returns], KEEPFILTERS ( Orders[Discount Band] = "Above 20%" ) )
VAR otherProfit = [Profit] - deepProfit
RETURN
    "Order lines discounted above 20% " & IF ( deepProfit < 0, "lost ", "earned " ) & """ + k("ABS ( deepProfit )") + """ & " on " & """ + k("deepSales") + """
        & " of sales. All other lines earned " & """ + k("otherProfit") + """ & ". Overall margin is " & FORMAT ( [% Margin], "0.0%" ) & "." """, None, "Insights", "Insight 4: discounts.")
M("Insight 5", """VAR y = [Selected Year]
VAR top20 = DIVIDE ( """ + PARETO.replace("__PCT__", "20").replace("VAR n", "VAR n").replace("\n", "\n    ") + """, [Total Sales] )
VAR segments = ADDCOLUMNS ( VALUES ( Customer[Segment] ), "@g", [Total Sales YoY %] )
VAR best = TOPN ( 1, FILTER ( segments, NOT ISBLANK ( [@g] ) ), [@g], DESC )
RETURN
    "The top 20% of customers bring " & FORMAT ( top20, "0.0%" ) & " of sales. Only " & FORMAT ( [New Customers], "#,0" ) & " of "
        & FORMAT ( [Customers], "#,0" ) & " customers were new in " & y & ", so growth came from existing customers."
        & IF ( COUNTROWS ( best ) > 0, " Fastest-growing segment: " & CONCATENATEX ( best, Customer[Segment], ", " ) & " (" & FORMAT ( MAXX ( best, [@g] ), "+0.0%;-0.0%" ) & ")." )""", None, "Insights", "Insight 5: customers.")
M("Insight Text", 'SWITCH ( SELECTEDVALUE ( Insight[No] ), 1, [Insight 1], 2, [Insight 2], 3, [Insight 3], 4, [Insight 4], 5, [Insight 5] )', None, "Insights", "The sentence for each row of the Insight table.")



# conditional-format colours (hex text, used in the format pane with "Field value")
G, R, D, GR, B = "#15803D", "#C2410C", "#111827", "#94A3B8", "#1D4ED8"
for m in ("Total Sales", "Profit", "Orders", "Customers"):
    M(f"Colour {m}", f'VAR g = [{m} YoY %]\nRETURN\n    IF ( ISBLANK ( g ), "{D}", IF ( g >= 0, "{G}", "{R}" ) )', None, "Colours", f"Green when {m} grew on last year, orange-red when it fell.")
M("Colour Margin", f'VAR ly = [% Margin LY]\nRETURN\n    IF ( ISBLANK ( ly ), "{D}", IF ( [% Margin] >= ly, "{G}", "{R}" ) )', None, "Colours", "Green when % margin is at or above last year.")
M("Colour Return Rate", f'VAR ly = [Return Rate LY]\nRETURN\n    IF ( ISBLANK ( ly ), "{D}", IF ( [Return Rate %] <= ly, "{G}", "{R}" ) )', None, "Colours", "Green when the return rate is at or below last year (lower is better).")
M("Colour Returned Sales", f'VAR ly = [Returned Sales LY]\nRETURN\n    IF ( ISBLANK ( ly ), "{D}", IF ( [Returned Sales] <= ly, "{G}", "{R}" ) )', None, "Colours", "Green when returned sales are at or below last year.")
M("Colour What-if", f'IF ( [Predicted Profit %] >= [% Margin], "{G}", "{R}" )', None, "Colours", "Green when the tested discounts give a higher profit % than today.")
M("Colour Loss", f'IF ( [Profit] < 0, "{R}", "{D}" )', None, "Colours", "Orange-red for a loss.")
M("Colour Discount", f'IF ( [% Avg Discount] > 0.2, "{R}", "{D}" )', None, "Colours", "Orange-red when the average discount is above 20% (such lines lose money).")
M("Colour Return Rate Bar", f'VAR overall = CALCULATE ( [Return Rate %], REMOVEFILTERS ( State ), REMOVEFILTERS ( Product ), REMOVEFILTERS ( Customer ) )\nRETURN\n    IF ( [Return Rate %] > overall, "{R}", "#64748B" )', None, "Colours", "Orange-red when a group returns more than the company as a whole.")
M("Colour Year", f'IF ( SELECTEDVALUE ( \'Trend Year\'[Year] ) = [Selected Year], "{B}", "{GR}" )', None, "Colours", "Blue for the selected year, grey for other years.")
M("Colour Year Return", f'IF ( SELECTEDVALUE ( \'Trend Year\'[Year] ) = [Selected Year], "{R}", "{GR}" )', None, "Colours", "Orange-red for the selected year, grey for other years.")
M("Colour Cohort", f'IF ( SELECTEDVALUE ( Customer[First Order Year] ) = [Selected Year], "{B}", "{GR}" )', None, "Colours", "Blue for customers who are new in the selected year.")
M("Colour Blue", f'"{B}"', None, "Colours", "Brand blue as a field value.")
M("Colour Tier", f'SWITCH ( SELECTEDVALUE ( Customer[Value Tier] ), "High", "{B}", "Medium", "#475569", "{GR}" )', None, "Colours", "Colour of a value tier.")
M("Colour Action", f'SWITCH ( SELECTEDVALUE ( Customer[Action] ), "Win back", "{R}", "Reward and keep", "{G}", "Grow", "{B}", "{D}" )', None, "Colours", "Colour of a marketing action.")
M("Colour Recent", f'IF ( SUM ( Customer[Sales Last 180 Days] ) = 0, "{R}", "{D}" )', None, "Colours", "Orange-red for customers who bought nothing in the last 180 days.")

# ---------------------------------------------------------------------------- TMDL writers

def indent(text, tabs):
    return "\n".join(("\t" * tabs + line) if line else "" for line in text.split("\n"))


def table_tmdl(t):
    name = t["name"]
    out = [f"/// {t['doc']}", f"table {q(name)}", f"\tlineageTag: {tag(name)}"]
    if name == "Date":
        out.append("\tdataCategory: Time")
    out.append("")
    for col, dtype, fmt, summ, desc, opt in t["cols"]:
        out += [f"\t/// {desc}", f"\tcolumn {q(col)}", f"\t\tdataType: {dtype}"]
        if opt.get("key"):
            out.append("\t\tisKey")
        if opt.get("hidden"):
            out.append("\t\tisHidden")
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        out += [f"\t\tlineageTag: {tag(name, col)}", f"\t\tsummarizeBy: {summ}", f"\t\tsourceColumn: {col}"]
        if opt.get("sortBy"):
            out.append(f"\t\tsortByColumn: {q(opt['sortBy'])}")
        out += ["", "\t\tannotation SummarizationSetBy = Automatic", ""]
        if dtype == "dateTime":
            out += ["\t\tannotation UnderlyingDateTimeDataType = Date", ""]
    out += [f"\tpartition {q(name)} = m", "\t\tmode: import", "\t\tsource =", indent(t["m"], 4), "", "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def measures_tmdl():
    out = ["/// All measures, in display folders.", "table _Measures", f"\tlineageTag: {tag('_Measures')}", ""]
    for name, dax, fmt, folder, desc in MEASURES:
        out.append(f"\t/// {desc}")
        if "\n" in dax:
            out += [f"\tmeasure {q(name)} =", indent(dax, 3)]
        else:
            out.append(f"\tmeasure {q(name)} = {dax}")
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        out += [f"\t\tdisplayFolder: {folder}", f"\t\tlineageTag: {tag('_Measures', name)}", ""]
    out += ["\tcolumn Dummy", "\t\tdataType: int64", "\t\tisHidden", "\t\tformatString: 0", f"\t\tlineageTag: {tag('_Measures', 'Dummy')}",
            "\t\tsummarizeBy: none", "\t\tsourceColumn: [Dummy]", "", "\t\tannotation SummarizationSetBy = Automatic", "",
            "\tpartition _Measures = calculated", "\t\tmode: import", '\t\tsource = ROW("Dummy", 0)', ""]
    return "\n".join(out)


def expressions_tmdl(source_file):
    out = ["/// Full path of the Excel workbook. Change it in Power BI Desktop: Transform data > Edit parameters.",
           f'expression SourceFile = "{source_file}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
           f"\tlineageTag: {tag('expr', 'SourceFile')}", "", "\tannotation PBI_ResultType = Text", ""]
    for name, doc, m, rtype in EXPRESSIONS:
        out += [f"/// {doc}", f"expression {name} =", indent(m, 2), f"\tlineageTag: {tag('expr', name)}", "",
                f"\tannotation PBI_ResultType = {rtype}", ""]
    return "\n".join(out)


def write(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def build_model(sm: Path, source_file=SOURCE_DEFAULT):
    d = sm / "definition"
    write(sm / "definition.pbism", json.dumps({"version": "4.2", "settings": {}}, indent=2) + "\n")
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1606\n\n")
    order = ["SourceFile"] + [e[0] for e in EXPRESSIONS] + [t["name"] for t in TABLES]
    write(d / "model.tmdl", "\n".join([
        "model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3", "\tdiscourageImplicitMeasures",
        "\tsourceQueryCulture: en-US", "\tdataAccessOptions", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
        "annotation __PBI_TimeIntelligenceEnabled = 0", "",
        "annotation PBI_QueryOrder = " + json.dumps(order, separators=(",", ":")), "",
        "ref table _Measures"] + [f"ref table {q(t['name'])}" for t in TABLES] + [""]))
    write(d / "expressions.tmdl", expressions_tmdl(source_file))
    rel = []
    for ft, fc, tt, tc in RELATIONSHIPS:
        rel += [f"relationship {tag('rel', ft, fc, tt)}", f"\tfromColumn: {q(ft)}.{q(fc)}", f"\ttoColumn: {q(tt)}.{q(tc)}", ""]
    write(d / "relationships.tmdl", "\n".join(rel))
    for t in TABLES:
        write(d / "tables" / f"{t['name']}.tmdl", table_tmdl(t))
    write(d / "tables" / "_Measures.tmdl", measures_tmdl())
