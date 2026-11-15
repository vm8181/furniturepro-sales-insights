"""Report pages for the FurniturePro report, written as PBIR files (one JSON file per visual)."""
import hashlib
import json
import shutil
from pathlib import Path

SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
BLUE, GREY, DARK, MID, GREEN, RED, AMBER = "#1D4ED8", "#9CA3AF", "#111827", "#374151", "#15803D", "#C2410C", "#D97706"
BORDER, PAGE_BG, WHITE = "#D0D5DD", "#EEF1F5", "#FFFFFF"
MS = "_Measures"
X0, X1, GAP = 10, 1270, 8          # left edge, right edge and gap between visuals
ROW_Y, ROW_H = 46, 58              # title and slicers
CARD_Y, CARD_H = 110, 80           # KPI cards
TOP, BOTTOM = 196, 712             # charts and tables


def hexid(*parts):
    return hashlib.sha1("/".join(parts).encode()).hexdigest()[:20]


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def txt(s):
    return lit("'" + s.replace("'", "''") + "'")


def color(c):
    return {"solid": {"color": lit(f"'{c}'")}}


def col_expr(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def mea_expr(prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": MS}}, "Property": prop}}


def color_by(measure):
    return {"solid": {"color": {"expr": mea_expr(measure)}}}


def C(entity, prop, display=None):
    p = {"field": col_expr(entity, prop), "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop}
    if display:
        p["displayName"] = display
    return p


def Mz(prop, display=None):
    p = {"field": mea_expr(prop), "queryRef": f"{MS}.{prop}", "nativeQueryRef": prop}
    if display:
        p["displayName"] = display
    return p


EACH_POINT = {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}


def box(title=None, sub=None, title_m=None, sub_m=None, pad=(4, 4, 8, 8), bg=WHITE, border=True, title_size=12):
    """Frame shared by every visual: white card, thin border, bold title, one-line explanation."""
    o = {"background": [{"properties": {"show": lit("true"), "color": color(bg), "transparency": lit("0D")}}] if bg else
         [{"properties": {"show": lit("false")}}],
         "border": [{"properties": {"show": lit("true" if border else "false"), "color": color(BORDER), "radius": lit("6D")}}],
         "dropShadow": [{"properties": {"show": lit("false")}}],
         "padding": [{"properties": {"top": lit(f"{pad[0]}D"), "bottom": lit(f"{pad[1]}D"), "left": lit(f"{pad[2]}D"), "right": lit(f"{pad[3]}D")}}]}
    if title or title_m:
        o["title"] = [{"properties": {"show": lit("true"), "text": {"expr": mea_expr(title_m)} if title_m else txt(title),
                                      "fontSize": lit(f"{title_size}D"), "fontColor": color(DARK), "bold": lit("true"),
                                      "fontFamily": txt("Segoe UI Semibold"), "titleWrap": lit("true")}}]
    else:
        o["title"] = [{"properties": {"show": lit("false")}}]
    if sub or sub_m:
        o["subTitle"] = [{"properties": {"show": lit("true"), "text": {"expr": mea_expr(sub_m)} if sub_m else txt(sub),
                                         "fontSize": lit("10D"), "fontColor": color(MID), "fontFamily": txt("Segoe UI"),
                                         "titleWrap": lit("true")}}]
    else:
        o["subTitle"] = [{"properties": {"show": lit("false")}}]
    return o


class Page:
    def __init__(self, key, name):
        self.key, self.name, self.id, self.visuals = key, name, hexid("page", key), []

    def add(self, key, vtype, x, y, w, h, query=None, objects=None, cobjects=None, extra=None):
        vid = hexid(self.key, key)
        z = (len(self.visuals) + 1) * 1000
        v = {"visualType": vtype}
        if query:
            v["query"] = query
        if objects:
            v["objects"] = objects
        if cobjects:
            v["visualContainerObjects"] = cobjects
        v["drillFilterOtherVisuals"] = True
        doc = {"$schema": f"{SCHEMA}/visualContainer/2.14.0/schema.json", "name": vid,
               "position": {"x": round(x, 2), "y": round(y, 2), "z": z, "width": round(w, 2), "height": round(h, 2), "tabOrder": z},
               "visual": v}
        if extra:
            v.update(extra)
        self.visuals.append((vid, doc))
        return vid


# ---------------------------------------------------------------------------- building blocks

def run(value, size, colour, bold=False):
    style = {"fontFamily": "Segoe UI Semibold" if bold else "Segoe UI", "fontSize": f"{size}pt", "color": colour}
    if bold:
        style["fontWeight"] = "bold"
    return {"textRuns": [{"value": value, "textStyle": style}]}


PAGES = [("insights", "Insights"), ("overview", "Overview"), ("targets", "Targets"), ("whatif", "What-if"),
         ("pareto", "Pareto"), ("returns", "Returns"), ("rfm", "RFM"), ("team", "Team")]


def header(pg):
    """Dark bar across the top with the report name and one button per page."""
    pg.add("header", "textbox", 0, 0, 1280, 40, objects={"general": [{"properties": {"paragraphs": [
        run("FurniturePro & Co.  ·  Sales and customer insights", 12, WHITE, True)]}}]},
        cobjects={"background": [{"properties": {"show": lit("true"), "color": color(DARK), "transparency": lit("0D")}}],
                  "border": [{"properties": {"show": lit("false")}}], "title": [{"properties": {"show": lit("false")}}],
                  "padding": [{"properties": {"top": lit("5D"), "bottom": lit("0D"), "left": lit("12D"), "right": lit("0D")}}]})

    def state(sid, fill, font, bold):
        return ({"properties": {"show": lit("true"), "fillColor": color(fill), "transparency": lit("0D")}, "selector": {"id": sid}},
                {"properties": {"show": lit("true"), "fontColor": color(font), "fontSize": lit("10D"), "bold": lit(bold),
                                "fontFamily": txt("Segoe UI Semibold"), "horizontalAlignment": txt("Center"),
                                "leftMargin": lit("2D"), "rightMargin": lit("2D")}, "selector": {"id": sid}},
                {"properties": {"show": lit("false")}, "selector": {"id": sid}})

    states = [state("default", DARK, "#E5E7EB", "false"), state("hover", MID, WHITE, "false"),
              state("press", MID, WHITE, "false"), state("selected", WHITE, DARK, "true")]
    pg.add("nav", "pageNavigator", 400, 5, 870, 30,
           objects={"fill": [s[0] for s in states], "text": [s[1] for s in states], "outline": [s[2] for s in states],
                    "pages": [{"properties": {"showHiddenPages": lit("false")}}]},
           cobjects={"background": [{"properties": {"show": lit("false")}}], "border": [{"properties": {"show": lit("false")}}],
                     "title": [{"properties": {"show": lit("false")}}]})


SLICERS = {  # key: entity, column, header, sync group, mode, single select, default value
    "year": ("Date", "Year", "Year", "year", "Dropdown", True, None),
    "region": ("State", "Region", "Region", "region", "Dropdown", False, None),
    "state": ("State", "State", "State", "state", "Dropdown", False, None),
    "city": ("Orders", "City", "City", None, "Dropdown", False, None),
    "segment": ("Customer", "Segment", "Segment", "segment", "Dropdown", False, None),
    "category": ("Product", "Category", "Category", "category", "Dropdown", False, None),
    "basis": ("Sales Basis", "Basis", "Sales basis", None, "Dropdown", True, "'After returns'"),
    "topn": ("Top N", "Top N", "Top N products", None, "Dropdown", True, "5L"),
    "tier": ("Customer", "Value Tier", "Value tier", None, "Dropdown", False, None),
    "action": ("Customer", "Action", "Marketing action", None, "Dropdown", False, None),
}


def slicer(pg, key, entity, column, head, x, y, w, h, mode="Dropdown", single=False, default=None, sync=None, strict=False):
    strict = strict or (single and default is not None)
    objects = {"data": [{"properties": {"mode": txt(mode)}}],
               "header": [{"properties": {"show": lit("true"), "text": txt(head), "fontColor": color(MID), "textSize": lit("10D"),
                                          "fontFamily": txt("Segoe UI Semibold")}}],
               "items": [{"properties": {"fontColor": color(DARK), "textSize": lit("11D")}}]}
    if default is not None:
        objects["selection"] = [{"properties": {"singleSelect": lit("true"), "strictSingleSelect": lit("true")}}]
    if mode == "Dropdown" and default is None:
        objects["selection"] = [{"properties": {"singleSelect": lit("true" if single else "false"),
                                                "strictSingleSelect": lit("true" if strict else "false")}}]
    if mode == "Single":
        objects["numericInputStyle"] = [{"properties": {"fontColor": color(DARK), "textSize": lit("12D")}}]
        objects["slider"] = [{"properties": {"show": lit("true"), "color": color(BLUE)}}]
    if default is not None:
        objects["general"] = [{"properties": {"filter": {"filter": {
            "Version": 2, "From": [{"Name": "s", "Entity": entity, "Type": 0}],
            "Where": [{"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": column}}],
                                            "Values": [[{"Literal": {"Value": default}}]]}}}]}}}}]
    extra = {"syncGroup": {"groupName": sync, "fieldChanges": True, "filterChanges": True}} if sync else None
    q = {"queryState": {"Values": {"projections": [{**C(entity, column), "active": True}]}}}
    pg.add("slicer-" + key, "slicer", x, y, w, h, q, objects, box(pad=(2, 2, 8, 8)), extra)


def frame(pg, title, desc, slicers, show_year=True, sw=146):
    """Header bar, page title on the left, slicers on the right."""
    header(pg)
    x = X1
    for key in reversed(slicers):
        entity, column, head, sync, mode, single, default = SLICERS[key]
        x -= sw
        slicer(pg, key, entity, column, head, x, ROW_Y, sw, ROW_H, mode, single, default, sync, strict=default is not None)
        x -= GAP
    if show_year:
        x -= 112
        pg.add("year-shown", "card", x, ROW_Y, 112, ROW_H,
               {"queryState": {"Values": {"projections": [Mz("Selected Year")]}}},
               {"labels": [{"properties": {"fontSize": lit("16D"), "color": color(BLUE), "fontFamily": txt("Segoe UI Semibold"),
                                           "labelDisplayUnits": lit("1D")}}],
                "categoryLabels": [{"properties": {"show": lit("false")}}]},
               box("Showing year", pad=(0, 0, 8, 8), title_size=10))
        x -= GAP
    pg.add("title", "textbox", X0, ROW_Y, x - X0, ROW_H, objects={"general": [{"properties": {"paragraphs": [
        run(title, 14, DARK, True), run(desc, 9, MID)]}}]},
        cobjects={"background": [{"properties": {"show": lit("false")}}], "border": [{"properties": {"show": lit("false")}}],
                  "title": [{"properties": {"show": lit("false")}}],
                  "padding": [{"properties": {"top": lit("0D"), "bottom": lit("0D"), "left": lit("2D"), "right": lit("2D")}}]})


def card(pg, key, measure, label, x, y, w, h, sub_m=None, sub=None, units="1D", precision=None, colour=DARK, size=22):
    labels = {"fontSize": lit(f"{size}D"), "color": color_by(colour) if colour.startswith("Colour") or colour == "Target Colour" else color(colour), "labelDisplayUnits": lit(units), "fontFamily": txt("Segoe UI Semibold")}
    if precision is not None:
        labels["labelPrecision"] = lit(f"{precision}L")
    pg.add("card-" + key, "card", x, y, w, h, {"queryState": {"Values": {"projections": [Mz(measure)]}}},
           {"labels": [{"properties": labels}], "categoryLabels": [{"properties": {"show": lit("false")}}]},
           {**box(label, sub, sub_m=sub_m, pad=(0, 0, 10, 10), title_size=10),
            "spacing": [{"properties": {"customizeSpacing": lit("true"), "spaceBelowTitle": lit("0D"), "spaceBelowSubTitle": lit("0D"),
                                        "spaceBelowTitleArea": lit("0D"), "verticalSpacing": lit("0D")}}]})


def cards(pg, specs, y=CARD_Y, h=CARD_H, x0=X0, x1=X1):
    """A row of equal KPI cards. spec: measure, label, small line (measure name, or text starting with '='), unit ('K' for thousands)."""
    w = (x1 - x0 - GAP * (len(specs) - 1)) / len(specs)
    for i, spec in enumerate(specs):
        measure, label, sub, unit = spec[:4]
        kw = {"units": "1000D", "precision": 1} if unit == "K" else {}
        if len(spec) > 4:
            kw["colour"] = spec[4]
        if sub and sub.startswith("="):
            kw["sub"] = sub[1:]
        elif sub:
            kw["sub_m"] = sub
        card(pg, measure, measure, label, x0 + i * (w + GAP), y, w, h, **kw)


def axis(show=True, **more):
    p = {"show": lit("true" if show else "false"), "fontSize": lit("10D"), "labelColor": color(MID), "showAxisTitle": lit("false")}
    p.update(more)
    return [{"properties": p}]


def chart(pg, key, vtype, x, y, w, h, title, sub, cat, ys, series=None, y2=None, sort=None, labels=True, units="0D",
          precision=None, value_axis=None, colour_m=None, legend=None, cat_axis=None, title_m=None, extra_objects=None, sub_m=None):
    """Bar, column, line and combo charts. ys / y2: list of (measure, display name, colour)."""
    y2 = y2 or []
    qs = {"Category": {"projections": [{**cat, "active": True}]}, "Y": {"projections": [Mz(m, d) for m, d, _ in ys]}}
    if y2:
        qs["Y2"] = {"projections": [Mz(m, d) for m, d, _ in y2]}
    if series:
        qs["Series"] = {"projections": [series]}
    query = {"queryState": qs}
    if sort:
        query["sortDefinition"] = {"sort": [{"field": f, "direction": d} for f, d in sort], "isDefaultSort": False}
    lab = {"show": lit("true" if labels else "false"), "fontSize": lit("10D"), "color": color(DARK), "labelDisplayUnits": lit(units)}
    if precision is not None:
        lab["labelPrecision"] = lit(f"{precision}L")
    is_bar = vtype in ("clusteredBarChart", "clusteredColumnChart")
    if value_axis is None:
        value_axis = not (is_bar and labels)
    multi = len(ys) + len(y2) > 1 or series is not None
    legend = multi if legend is None else legend
    objects = {
        "labels": [{"properties": lab}],
        "valueAxis": axis(value_axis, gridlineStyle=txt("dotted"), labelDisplayUnits=lit(units)),
        "categoryAxis": axis(True, **(cat_axis or {})),
        "legend": [{"properties": {"show": lit("true" if legend else "false"), "position": txt("Top"), "fontSize": lit("10D"),
                                   "labelColor": color(DARK), "showTitle": lit("false")}}],
        "general": [{"properties": {"responsive": lit("false")}}],
    }
    if len(ys) == 1 and not y2 and not series:
        points = [{"properties": {"fill": color(ys[0][2])}}] if ys[0][2] else []
    else:
        points = [{"properties": {"fill": color(c)}, "selector": {"metadata": f"{MS}.{m}"}} for m, _, c in ys + y2 if c]
    if colour_m:
        points = [{"properties": {"fill": color_by(colour_m)}, "selector": EACH_POINT}]
    if points:
        objects["dataPoint"] = points
    if y2:
        objects["labels"].append({"properties": {"show": lit("false")}, "selector": {"metadata": f"{MS}.{y2[0][0]}"}})
    if vtype in ("lineChart", "lineClusteredColumnComboChart"):
        objects["lineStyles"] = [{"properties": {"strokeWidth": lit("2D"), "showMarker": lit("true"), "markerSize": lit("4D")}}]
    if extra_objects:
        for name, entries in extra_objects.items():
            objects[name] = objects.get(name, []) + entries if name in ("lineStyles", "labels", "dataPoint") else entries
    return pg.add(key, vtype, x, y, w, h, query, objects, box(title, sub, title_m=title_m, sub_m=sub_m))


def table(pg, key, x, y, w, h, title, sub, cols, sort=None, totals=False, font=10, pad=3, extra_objects=None, vtype="tableEx", sub_m=None, title_m=None):
    """cols: list of (projection, width)."""
    query = {"queryState": {"Values": {"projections": [p for p, _ in cols]}}}
    if sort:
        query["sortDefinition"] = {"sort": [{"field": f, "direction": d} for f, d in sort], "isDefaultSort": False}
    objects = {"values": [{"properties": {"fontSize": lit(f"{font}D"), "fontColor": color(DARK), "wordWrap": lit("true")}}],
               "columnHeaders": [{"properties": {"fontSize": lit(f"{font}D"), "fontColor": color(DARK), "bold": lit("true"),
                                                 "backColor": color("#E8ECF2"), "wordWrap": lit("true"), "autoSizeColumnWidth": lit("false")}}],
               "grid": [{"properties": {"gridVertical": lit("false"), "rowPadding": lit(f"{pad}D")}}],
               "total": [{"properties": {"totals": lit("true" if totals else "false"), "fontSize": lit(f"{font}D"), "bold": lit("true")}}],
               "columnWidth": [{"properties": {"value": lit(f"{float(wd)}D")}, "selector": {"metadata": p["queryRef"]}} for p, wd in cols]}
    if extra_objects:
        for name, entries in extra_objects.items():
            objects[name] = objects.get(name, []) + entries
    return pg.add(key, vtype, x, y, w, h, query, objects, box(title, sub, sub_m=sub_m, title_m=title_m))


def cf(measure, ref):
    """Font colour of one table column from a colour measure (conditional formatting by field value)."""
    return {"properties": {"fontColor": color_by(measure)}, "selector": {**EACH_POINT, "metadata": ref}}


def ref(p):
    return p["queryRef"]


def by(prop, direction="Descending"):
    return (mea_expr(prop), direction)


def by_col(entity, prop, direction="Ascending"):
    return (col_expr(entity, prop), direction)


MONTH = C("Date", "Month")
MONTH_SORT = [by_col("Date", "Month")]
CY_LY = [("Total Sales", "Selected year", BLUE), ("Total Sales LY", "Year before", GREY)]


def split_bar(pg, key, x, y, w, h, title, cat, sub="Total sales after returns, selected year against the year before."):
    chart(pg, key, "clusteredBarChart", x, y, w, h, title, sub, cat, CY_LY, sort=[by("Total Sales")], units="1000D", precision=1)


def rate_bar(pg, key, x, y, w, h, title, cat, sub="Returned sales as a % of sales before returns. Orange-red = above the company rate."):
    chart(pg, key, "clusteredBarChart", x, y, w, h, title, sub, cat, [("Return Rate %", "Return rate", RED)], sort=[by("Return Rate %")],
          colour_m="Colour Return Rate Bar")


def discount_profit(pg, key, x, y, w, h, sub):
    chart(pg, key, "clusteredColumnChart", x, y, w, h, "Profit by discount given", sub, C("Orders", "Discount"),
          [("Profit", "Profit", None)], sort=[by_col("Orders", "Discount")], units="1000D", precision=1, colour_m="Profit Colour",
          cat_axis={"axisType": txt("Categorical")})


# ---------------------------------------------------------------------------- pages

def page_insights():
    pg = Page("insights", "Insights")
    frame(pg, "Key insights for the year-end meeting", "Five things the data says. Every sentence follows the slicers.",
          ["year", "region", "segment", "category"])
    cards(pg, [("Total Sales", "Total sales (after returns)", "Total Sales YoY Text", "K", "Colour Total Sales"),
               ("% Target Achieved", "% of target achieved", "Target Text", None, "Target Colour"),
               ("% Margin", "% margin (profit / sales)", "% Margin Text", None, "Colour Margin"),
               ("Profit", "Profit", "Profit YoY Text", "K", "Colour Profit"),
               ("Return Rate %", "Return rate", "Return Rate Text", None, "Colour Return Rate"),
               ("New Customers", "New customers", "New Customers Text", None)])
    table(pg, "insights", X0, TOP, 560, BOTTOM - TOP, "Top 5 insights", "Read top to bottom. The pages in the bar above show the detail.",
          [(C("Insight", "No", "#"), 30), (C("Insight", "Theme", "Theme"), 118), (Mz("Insight Text", "What the data says"), 394)],
          sort=[by_col("Insight", "No")], font=11, pad=7)
    x, w, h = X0 + 560 + GAP, (X1 - X0 - 560 - 2 * GAP) / 2, (BOTTOM - TOP - GAP) / 2
    chart(pg, "month-target", "lineClusteredColumnComboChart", x, TOP, w, h, "Sales against target, by month",
          "Columns: sales (green = month on target). Line: monthly target.", MONTH, [("Sales For Target", "Sales", None)],
          y2=[("Target", "Target", DARK)], sort=MONTH_SORT, labels=False, units="1000D", colour_m="Target Colour", value_axis=True, legend=False)
    rate_bar(pg, "returns-region", x + w + GAP, TOP, w, h, "Return rate by region", C("State", "Region"), "Orange-red = above the company rate.")
    discount_profit(pg, "discount-profit", x, TOP + h + GAP, w, h, "Profit at each discount rate. Orange-red = a loss.")
    split_bar(pg, "segment", x + w + GAP, TOP + h + GAP, w, h, "Total sales by segment", C("Customer", "Segment"),
              "Blue = selected year, grey = year before.")
    return pg


def page_overview():
    pg = Page("overview", "Overview")
    frame(pg, "Sales performance: overall", "Total sales = sales minus returns. Trend charts always show 4 years.",
          ["year", "region", "segment", "category"])
    cards(pg, [("Total Sales", "Total sales (after returns)", "Total Sales YoY Text", "K", "Colour Total Sales"),
               ("% Target Achieved", "% of target achieved", "Target Text", None, "Target Colour"),
               ("% Margin", "% margin (profit / sales)", "% Margin Text", None, "Colour Margin"),
               ("Profit", "Profit", "Profit YoY Text", "K", "Colour Profit"),
               ("Orders", "Orders", "Orders YoY Text", None, "Colour Orders"),
               ("Customers", "Customers", "Customers YoY Text", None, "Colour Customers")])
    h1 = 226
    chart(pg, "years", "clusteredColumnChart", X0, TOP, 300, h1, "Total sales by year", "Selected year in blue.",
          C("Trend Year", "Year"), [("Sales Trend", "Total sales", BLUE)], sort=[by_col("Trend Year", "Year")], units="1000D", precision=1,
          colour_m="Colour Year", cat_axis={"axisType": txt("Categorical")})
    chart(pg, "months", "lineChart", X0 + 300 + GAP, TOP, X1 - X0 - 300 - GAP, h1, "Total sales by month, last 4 years",
          "One line per year. The latest year is blue.",
          MONTH, [("Sales Trend", "Total sales", None)], series=C("Trend Year", "Year"), sort=MONTH_SORT, labels=False, units="1000D")
    y2, h2 = TOP + h1 + GAP, BOTTOM - TOP - h1 - GAP
    w = (X1 - X0 - 2 * GAP) / 3
    sub = "Blue = selected year, grey = year before."
    split_bar(pg, "segment", X0, y2, w, h2, "Total sales by segment", C("Customer", "Segment"), sub)
    split_bar(pg, "category", X0 + w + GAP, y2, w, h2, "Total sales by category", C("Product", "Category"), sub)
    split_bar(pg, "region", X0 + 2 * (w + GAP), y2, w, h2, "Total sales by region", C("State", "Region"), sub)
    return pg


def page_targets():
    pg = Page("targets", "Targets")
    frame(pg, "Target achieved or not achieved", "Green = target reached. Orange-red = target missed.",
          ["year", "region", "state", "basis"], sw=150)
    cards(pg, [("Sales For Target", "Sales compared with target", "Basis Text", "K"),
               ("Target", "Target", "Months Achieved Text", "K"),
               ("% Target Achieved", "% of target achieved", "Target Text", None, "Target Colour"),
               ("Regions Achieved Text", "Regions on target", "=Regions that reached the target", None),
               ("States Achieved Text", "States on target", "States Missed Text", None)])
    h1 = 214
    chart(pg, "regions", "clusteredBarChart", X0, TOP, 330, h1, "% of target achieved by region", "100% = target reached.",
          C("State", "Region"), [("% Target Achieved", "% of target", None)], sort=[by("% Target Achieved")], colour_m="Target Colour")
    chart(pg, "months", "lineClusteredColumnComboChart", X0 + 330 + GAP, TOP, X1 - X0 - 330 - GAP, h1, "Sales against target, by month",
          "Columns: sales (green = month reached its target). Line: monthly target.", MONTH,
          [("Sales For Target", "Sales", None)], y2=[("Target", "Target", DARK)], sort=MONTH_SORT, units="1000D", precision=1,
          colour_m="Target Colour", value_axis=True, legend=False)
    y2, h2 = TOP + h1 + GAP, BOTTOM - TOP - h1 - GAP
    cw = 640
    chart(pg, "states", "clusteredColumnChart", X0, y2, cw, h2, "% of target achieved by state",
          "Best first. The axis stops at 200%: small targets go far above it.", C("State", "State"),
          [("% Target Achieved", "% of target", None)], sort=[by("% Target Achieved")], labels=False, colour_m="Target Colour", value_axis=True,
          extra_objects={"valueAxis": axis(True, gridlineStyle=txt("dotted"), end=lit("2D"), start=lit("0D")),
                         "categoryAxis": axis(True, fontSize=lit("8D"), preferredCategoryWidth=lit("8D"))})
    cols = [(C("State", "State", "State"), 96), (C("State", "Region", "Region"), 54), (C("State", "Region Head", "Region head"), 98),
            (Mz("Sales For Target", "Sales"), 58), (Mz("Target", "Target"), 58), (Mz("Target Gap", "Gap"), 60),
            (Mz("% Target Achieved", "% target"), 52), (Mz("Target Status", "Status"), 70)]
    table(pg, "state-table", X0 + cw + GAP, y2, X1 - X0 - cw - GAP, h2, "States, biggest gap first", "Gap = sales minus target.",
          cols, sort=[by("Target Gap", "Ascending")], font=9, pad=2,
          extra_objects={"values": [cf("Target Colour", ref(cols[i][0])) for i in (5, 6, 7)]})
    return pg


def page_whatif():
    pg = Page("whatif", "What-if")
    frame(pg, "What-if: change the discount, see the profit %", "Pick a discount for each segment. List prices and costs stay the same.",
          ["year", "region", "category"])
    w = (X1 - X0 - 3 * GAP) / 4
    for i, seg in enumerate(("Consumer", "Corporate", "Home Office")):
        slicer(pg, "disc-" + seg, f"Discount {seg}", f"{seg} Discount %", f"Your input: {seg} discount %", X0 + i * (w + GAP), CARD_Y, w, 58,
               mode="Dropdown", single=True, default="20L")
    slicer(pg, "volume", "Volume Response", "Extra Units %", "Optional: % more units per 1 point of discount", X0 + 3 * (w + GAP), CARD_Y, w, 58,
           mode="Dropdown", single=True, default="0D")
    y = CARD_Y + 58 + GAP
    cards(pg, [("% Margin", "Current profit %", "Current Profit Text", None, "Colour Margin"),
               ("Predicted Profit %", "Predicted profit %", "=At the discounts you picked", None, "Colour What-if"),
               ("Profit % Change", "Change in profit % (points)", "=Predicted minus current", None, "Colour What-if"),
               ("Predicted Profit", "Predicted profit", "Predicted Profit Text", "K", "Colour What-if"),
               ("Predicted Sales", "Predicted sales", "Predicted Sales Text", "K"),
               ("Tested Discount %", "Discount tested (all segments)", "Tested Discount Text", None)], y=y)
    top = y + CARD_H + GAP
    lw, th = 640, 190
    cols = [(C("Customer", "Segment", "Segment"), 92), (Mz("% Discount Given", "Disc. now"), 62), (Mz("Tested Discount %", "Disc. test"), 62),
            (Mz("Sales Before Returns", "Sales now"), 72), (Mz("Predicted Sales", "Sales new"), 72), (Mz("Profit", "Profit now"), 66),
            (Mz("Predicted Profit", "Profit new"), 66), (Mz("% Margin", "% now"), 52), (Mz("Predicted Profit %", "% new"), 52)]
    table(pg, "segments", X0, top, lw, th, "Current against predicted, by segment", "Sales are before returns (profit % = profit / sales amount).",
          cols, sort=[by("Sales Before Returns")], totals=True, pad=4, font=10,
          extra_objects={"values": [cf("Colour What-if", ref(cols[8][0]))]})
    chart(pg, "segments-chart", "clusteredColumnChart", X0, top + th + GAP, lw, BOTTOM - top - th - GAP, "Profit % by segment: current against predicted",
          "Grey = today. Blue = at the discounts you picked.", C("Customer", "Segment"),
          [("% Margin", "Current profit %", GREY), ("Predicted Profit %", "Predicted profit %", BLUE)], sort=[by("Sales Before Returns")])
    discount_profit(pg, "discount-profit", X0 + lw + GAP, top, X1 - X0 - lw - GAP, BOTTOM - top,
                    "Why discount matters: profit of order lines at each discount given today. Lines above 20% lose money (orange-red).")
    return pg


def page_pareto():
    pg = Page("pareto", "Pareto")
    frame(pg, "Pareto: how much do the top customers bring in?", "Pick a top % of customers. Ranked by sales in the year.",
          ["year", "region", "segment", "category"])
    slicer(pg, "top-pct", "Top Customer %", "Top %", "Your input: top % of customers", X0, CARD_Y, 300, CARD_H, mode="Dropdown", single=True, default="20L")
    cards(pg, [("Top Customers Share %", "Share of total sales from them", "Top Customers Count Text", None, "Colour Blue"),
               ("Top Customers Sales", "Their total sales", "Top Customers Sales Text", "K"),
               ("Customers For 80%", "Customers that make 80% of sales", "Customers For 80% Text", None),
               ("Sales Per Customer", "Average sales per customer", "Customers YoY Text", None)], x0=X0 + 300 + GAP)
    lw = 740
    chart(pg, "curve", "lineChart", X0, TOP, lw, BOTTOM - TOP, "Share of total sales by top % of customers",
          "The top x% of customers bring y% of sales. The orange-red dot is your input.", C("Pareto Axis", "Top % of Customers"),
          [("Pareto Curve %", "Share of total sales", BLUE), ("Pareto Selected Point", "Your input", RED)], labels=False, legend=True,
          value_axis=True, cat_axis={"showAxisTitle": lit("true"), "titleText": txt("Top % of customers"), "titleFontSize": lit("10D")},
          extra_objects={"lineStyles": [{"properties": {"showMarker": lit("false")}, "selector": {"metadata": f"{MS}.Pareto Curve %"}},
                                        {"properties": {"markerSize": lit("9D")}, "selector": {"metadata": f"{MS}.Pareto Selected Point"}}],
                         "labels": [{"properties": {"show": lit("true"), "fontSize": lit("11D"), "color": color(DARK)},
                                     "selector": {"metadata": f"{MS}.Pareto Selected Point"}}]})
    table(pg, "customers", X0 + lw + GAP, TOP, X1 - X0 - lw - GAP, BOTTOM - TOP, "Customers inside the chosen top %",
          "Biggest first. The list follows your input.",
          [(Mz("Top Customer Rank", "#"), 36), (C("Customer", "Customer Name", "Customer"), 140), (C("Customer", "Segment", "Segment"), 90),
           (Mz("Top Customer Sales", "Total sales"), 76), (Mz("Top Customer Share", "% total"), 62), (Mz("Top Customer Orders", "Orders"), 60)],
          sort=[by("Top Customer Sales")], pad=2)
    return pg


def page_returns():
    pg = Page("returns", "Returns")
    frame(pg, "Returns: what comes back, and from where", "A returned order removes all its lines from total sales.",
          ["year", "region", "category", "topn"])
    cards(pg, [("Returned Sales", "Returned sales", "Returned Sales Text", "K", "Colour Returned Sales"),
               ("Return Rate %", "Return rate (% of sales)", "Return Rate Text", None, "Colour Return Rate"),
               ("Returned Orders", "Returned orders", "Returned Orders Text", None),
               ("Returned Quantity", "Returned units", "Returned Quantity Text", None),
               ("Returned Profit", "Profit on returned orders", "=Profit booked, then lost", "K"),
               ("Total Sales", "Total sales after returns", "Total Sales YoY Text", "K", "Colour Total Sales")])
    lw, mw = 440, 480
    chart(pg, "top-products", "clusteredBarChart", X0, TOP, lw, BOTTOM - TOP, None, "Ranked by returned sales. Change the count with Top N.",
          C("Product", "Product Name"), [("Returned Sales Top N", "Returned sales", RED)], sort=[by("Returned Sales Top N")], units="1000D", precision=1,
          title_m="Top N Title", cat_axis={"maxMarginFactor": lit("50L")})
    x = X0 + lw + GAP
    cols = [(Mz("Return Rank Top N", "#"), 28), (C("Product", "Product Name", "Product"), 190), (Mz("Returned Category", "Category"), 76),
            (Mz("Top Return Region", "Region"), 52), (Mz("Returned Sales Top N", "Returned sales"), 68), (Mz("Returned Quantity Top N", "Units"), 40)]
    table(pg, "product-table", x, TOP, mw, 300, "Category and region of the most-returned products", "Region = where most of its returns happened.",
          cols, sort=[by("Returned Sales Top N")], font=9, pad=2)
    chart(pg, "rate-years", "clusteredColumnChart", x, TOP + 300 + GAP, mw, BOTTOM - TOP - 300 - GAP, "Return rate by year", "Selected year in orange-red.",
          C("Trend Year", "Year"), [("Return Rate Trend", "Return rate", RED)], sort=[by_col("Trend Year", "Year")], colour_m="Colour Year Return",
          cat_axis={"axisType": txt("Categorical")})
    x += mw + GAP
    w, h = X1 - x, (BOTTOM - TOP - 2 * GAP) / 3
    rate_bar(pg, "rate-region", x, TOP, w, h, "Return rate by region", C("State", "Region"), "Orange-red = above company rate.")
    rate_bar(pg, "rate-category", x, TOP + h + GAP, w, h, "Return rate by category", C("Product", "Category"), "Orange-red = above company rate.")
    rate_bar(pg, "rate-segment", x, TOP + 2 * (h + GAP), w, h, "Return rate by segment", C("Customer", "Segment"), "Orange-red = above company rate.")
    return pg


def tier_colours():
    out = []
    for tier, c in (("High", BLUE), ("Medium", "#64748B"), ("Low", "#CBD5E1")):
        out.append({"properties": {"fill": color(c)}, "selector": {"data": [{"scopeId": {"Comparison": {
            "ComparisonKind": 0, "Left": col_expr("Customer", "Value Tier"), "Right": {"Literal": {"Value": f"'{tier}'"}}}}}]}})
    return out


def page_rfm():
    pg = Page("rfm", "RFM")
    frame(pg, "Customer value: RFM", "All years. R = sales in last 180 days, F = orders, M = total sales; each scored 1 to 5.",
          ["segment", "tier", "action"], show_year=False, sw=160)
    cards(pg, [("RFM Customers", "Customers", "RFM Customers Text", None),
               ("RFM Sales", "Total sales", "RFM Sales Text", "K"),
               ("RFM Sales Last 180 Days", "Sales in the last 180 days", "RFM 180 Text", "K"),
               ("High Value Share %", "Sales from high value customers", "High Value Text", None, "Colour Blue"),
               ("Win Back Customers", "Customers to win back", "=Big spenders, nothing in 180 days", None, "#C2410C"),
               ("Win Back Sales", "Sales of win-back customers", "Win Back Text", "K", "#C2410C")])
    h1 = 262
    q = {"queryState": {"Category": {"projections": [{**C("Customer", "Customer Name"), "active": True}]},
                        "X": {"projections": [Mz("RFM Orders", "Orders")]}, "Y": {"projections": [Mz("RFM Sales", "Total sales")]},
                        "Series": {"projections": [C("Customer", "Value Tier")]}}}
    pg.add("scatter", "scatterChart", X0, TOP, 400, h1, q,
           {"categoryAxis": axis(True, showAxisTitle=lit("true"), titleText=txt("Number of orders"), titleFontSize=lit("10D")),
            "valueAxis": axis(True, showAxisTitle=lit("true"), titleText=txt("Total sales"), titleFontSize=lit("10D"), gridlineStyle=txt("dotted")),
            "legend": [{"properties": {"show": lit("true"), "position": txt("Top"), "fontSize": lit("10D"), "labelColor": color(DARK), "showTitle": lit("false")}}],
            "dataPoint": tier_colours(), "bubbles": [{"properties": {"bubbleSize": lit("-25L")}}]},
           box("Every customer: orders against sales", "One dot per customer, coloured by value tier."))
    tcols = [(C("Customer", "Value Tier", "Tier"), 56), (Mz("RFM Customers", "Customers"), 66), (Mz("RFM % of Customers", "% cust."), 52),
             (Mz("RFM Sales", "Total sales"), 76), (Mz("RFM % of Sales", "% sales"), 52), (Mz("RFM Avg Orders", "Avg orders"), 60),
             (Mz("RFM Sales Last 180 Days", "Last 180d"), 66)]
    table(pg, "tiers", X0 + 400 + GAP, TOP, 450, h1, "Low, medium and high value customers",
          "High = score 12 to 15, Medium = 8 to 11, Low = 3 to 7.",
          tcols, sort=[by_col("Customer", "Value Tier")], totals=True, pad=8, font=10,
          extra_objects={"values": [cf("Colour Tier", ref(tcols[0][0]))]})
    x = X0 + 400 + 450 + 2 * GAP
    heat = {"queryState": {"Rows": {"projections": [{**C("Customer", "F Score", "F score"), "active": True}]},
                           "Columns": {"projections": [{**C("Customer", "M Score", "M score"), "active": True}]},
                           "Values": {"projections": [Mz("RFM Customers", "Customers")]}}}
    pg.add("heat", "pivotTable", x, TOP, X1 - x, h1, heat,
           {"values": [{"properties": {"fontSize": lit("11D")}},
                       {"properties": {"backColor": color_by("Heat Colour"), "fontColor": color_by("Heat Font")},
                        "selector": {**EACH_POINT, "metadata": f"{MS}.RFM Customers"}}],
            "columnHeaders": [{"properties": {"fontSize": lit("10D"), "bold": lit("true"), "fontColor": color(DARK)}}],
            "rowHeaders": [{"properties": {"fontSize": lit("10D"), "bold": lit("true"), "fontColor": color(DARK)}}],
            "subTotals": [{"properties": {"rowSubtotals": lit("false"), "columnSubtotals": lit("false")}}],
            "grid": [{"properties": {"rowPadding": lit("5D")}}]},
           box("Score heat map: customers", "Rows = F score, columns = M score. Darker = more."))
    y2, h2 = TOP + h1 + GAP, BOTTOM - TOP - h1 - GAP
    acols = [(C("Customer", "Action", "Action"), 112), (C("Customer", "Action Rule", "Rule"), 172), (Mz("RFM Customers", "Customers"), 70),
             (Mz("RFM Sales", "Total sales"), 76), (Mz("RFM % of Sales", "% sales"), 56), (C("Customer", "Action Advice", "What to do"), 180)]
    table(pg, "actions", X0, y2, 690, h2, "Who to target: one marketing action per customer",
          "Bonus scoring for targeted marketing. Click a row to list its customers.",
          acols, sort=[by_col("Customer", "Action")], totals=True, pad=3, font=9,
          extra_objects={"values": [cf("Colour Action", ref(acols[0][0]))]})
    lcols = [(C("Customer", "Customer Name", "Customer"), 118), (C("Customer", "Segment", "Segment"), 82), (C("Customer", "Value Tier", "Tier"), 56),
             (C("Customer", "RFM Score", "RFM"), 40), (C("Customer", "Lifetime Orders", "Orders"), 52), (C("Customer", "Lifetime Sales", "Total sales"), 66),
             (C("Customer", "Sales Last 180 Days", "Last 180d"), 62), (C("Customer", "Action", "Action"), 82)]
    table(pg, "customer-list", X0 + 690 + GAP, y2, X1 - X0 - 690 - GAP, h2, "Customer list", "Follows the slicers and any row you click.",
          lcols, sort=[by_col("Customer", "Lifetime Sales", "Descending")], font=9, pad=2,
          extra_objects={"values": [cf("Colour Tier", ref(lcols[2][0])), cf("Colour Recent", ref(lcols[6][0])), cf("Colour Action", ref(lcols[7][0]))]})
    return pg


def page_team():
    pg = Page("team", "Team")
    frame(pg, "New customers and sales team", "New = first order ever is in the year. State and City filter the team.",
          ["year", "segment", "region", "state", "city"], sw=134)
    cards(pg, [("New Customers", "New customers", "New Customers Text", None, "Colour Blue"),
               ("New Customer Sales", "Total sales from new customers", "New Customer Sales Text", "K"),
               ("New Customer Quantity", "Quantity bought by new customers", "New Customer Quantity Text", None),
               ("New Customer Discount Amount", "Discount amount to new customers", "New Customer Discount Text", None),
               ("Retention %", "Last year's customers who came back", "Retention Text", None),
               ("Sales Per Customer", "Average sales per customer", "Sales Per Customer Text", None)])
    lw, ch = 480, 180
    chart(pg, "cohort", "clusteredColumnChart", X0, TOP, lw, ch, "Customers by the year of their first order",
          "Blue = new in the selected year. Most buyers were won earlier.", C("Customer", "First Order Year"),
          [("Customers", "Customers", BLUE)], sort=[by_col("Customer", "First Order Year")], colour_m="Colour Cohort", cat_axis={"axisType": txt("Categorical")})
    table(pg, "new-customers", X0, TOP + ch + GAP, lw, BOTTOM - TOP - ch - GAP, "New customers in the selected year", "Total sales are after returns.",
          [(C("Customer", "Customer Name", "Customer"), 118), (C("Customer", "Segment", "Segment"), 80), (C("Customer", "First Order Date", "First order"), 86),
           (Mz("New Customer Sales", "Total sales"), 66), (Mz("New Customer Quantity", "Qty"), 40), (Mz("New Customer Discount Amount", "Discount"), 60)],
          sort=[by("New Customer Sales")], totals=True, pad=2, font=9)
    tcols = [(C("Customer", "Segment", "Segment"), 104), (Mz("Top 5 Rank", "Rank"), 44), (C("Sales Person", "Sales Person", "Sales person"), 180),
             (Mz("Top 5 Sales", "Total sales"), 100), (Mz("Top 5 Avg Discount", "% avg discount"), 100), (Mz("Top 5 Margin", "% margin"), 80),
             (Mz("Top 5 Order Lines", "Lines"), 60)]
    table(pg, "top-sales-people", X0 + lw + GAP, TOP, X1 - X0 - lw - GAP, BOTTOM - TOP, "Top 5 sales people in each segment",
          "By total sales after returns. Discount above 20% and negative margins show in orange-red.",
          tcols, sort=[by_col("Customer", "Segment"), by("Top 5 Sales")], font=10, pad=3,
          extra_objects={"values": [cf("Colour Discount", ref(tcols[4][0])), cf("Colour Loss", ref(tcols[5][0]))]})
    return pg


THEME = {
    "name": "FurniturePro",
    "dataColors": ["#CBD5E1", "#94A3B8", "#475569", BLUE, AMBER, "#0F766E", "#7C3AED", RED],
    "background": WHITE, "foreground": DARK, "tableAccent": BLUE, "good": GREEN, "neutral": AMBER, "bad": RED,
}


def build_report(rp: Path, name: str, base_theme: Path):
    rd = rp / "definition"
    if rd.exists():
        shutil.rmtree(rd)

    def wj(path, obj):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(obj, indent=2, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")

    pages = [page_insights(), page_overview(), page_targets(), page_whatif(), page_pareto(), page_returns(), page_rfm(), page_team()]
    wj(rp / "definition.pbir", {"version": "4.0", "datasetReference": {"byPath": {"path": f"../{name}.SemanticModel"}}})
    wj(rd / "version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    ver = {"visual": "2.14.0", "report": "3.4.0", "page": "2.3.1"}
    wj(rd / "report.json", {
        "$schema": f"{SCHEMA}/report/3.3.0/schema.json",
        "themeCollection": {"baseTheme": {"name": "Fluent2-CY26SU10", "reportVersionAtImport": ver, "type": "SharedResources"},
                            "customTheme": {"name": "FurnitureProTheme.json", "reportVersionAtImport": ver, "type": "RegisteredResources"}},
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": "Fluent2-CY26SU10", "path": "BaseThemes/Fluent2-CY26SU10.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": "FurnitureProTheme.json", "path": "FurnitureProTheme.json", "type": "CustomTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized", "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True, "useDefaultAggregateDisplayName": True}})
    wj(rp / "StaticResources" / "RegisteredResources" / "FurnitureProTheme.json", THEME)
    dest = rp / "StaticResources" / "SharedResources" / "BaseThemes" / "Fluent2-CY26SU10.json"
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(base_theme, dest)
    wj(rd / "pages" / "pages.json", {"$schema": f"{SCHEMA}/pagesMetadata/1.1.0/schema.json", "pageOrder": [p.id for p in pages],
                                     "activePageName": pages[0].id})
    for p in pages:
        wj(rd / "pages" / p.id / "page.json", {
            "$schema": f"{SCHEMA}/page/2.1.0/schema.json", "name": p.id, "displayName": p.name, "displayOption": "FitToPage",
            "height": 720, "width": 1280,
            "objects": {"background": [{"properties": {"color": color(PAGE_BG), "transparency": lit("0D")}}],
                        "outspace": [{"properties": {"color": color(PAGE_BG), "transparency": lit("0D")}}]}})
        for vid, doc in p.visuals:
            wj(rd / "pages" / p.id / "visuals" / vid / "visual.json", doc)
    return pages
