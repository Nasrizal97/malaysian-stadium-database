# Malaysian Stadium Database

A menu-driven, macro-enabled **Microsoft Excel database** of 50 sports stadiums and arenas in Malaysia, with a data entry form, search, PivotTables, charts and slicers. Built for **IMC411 Database Management System** (Bachelor of Information Science (Hons) Information Systems Management, UiTM).

## What this project shows

|Skill|How it is used|
|-|-|
|Data modelling|18 attributes per stadium, with a unique `Stadium ID` (e.g. `STD-001`) and audit fields (`Created By`, `Created On`)|
|Data types|Text, integer, currency (RM), date and hyperlink, each formatted correctly|
|Data entry \& validation|Excel form with required-field markers, dropdown lists, number/date rules and VBA Save / Reset|
|Search|Single search box that filters across all columns using the `FILTER` function|
|Analysis|4 PivotTables, 4 PivotCharts, 4 slicers and KPI cards on an Analysis sheet|
|Reporting|Written report covering scope, data management, evidence of work, charts and conclusion|

## Dataset scope

* **Domain:** sports stadiums and arenas in Malaysia (outdoor stadiums, indoor arenas, hockey stadiums, aquatic centres)
* **Size filter:** capacity of at least 3,000 seats
* **Coverage:** 50 venues across all 13 states and Kuala Lumpur
* **Opening years:** 1954 to 2026

A plain-text copy of the data is in [`data/stadiums.csv`](data/stadiums.csv).

> \*\*Data quality note:\*\* names, states, capacity and operators were collected from public sources. Construction cost, floodlight lux and last renovation date are published inconsistently, so treat those three fields as approximate and verify them against official sources before relying on them.

## Data dictionary

|#|Field|Type|Example|
|-|-|-|-|
|1|ID|Integer|1|
|2|Stadium ID|Text|STD-001|
|3|Stadium Name|Text|Bukit Jalil National Stadium|
|4|State|Text (list)|Kuala Lumpur|
|5|City|Text|Bukit Jalil|
|6|Capacity|Integer|87,411|
|7|Year Opened|Integer|1998|
|8|Venue Type|Text (list)|Outdoor|
|9|Surface Type|Text (list)|Zeon Zoysia|
|10|Home Tenant|Text|Malaysia National Football Team|
|11|Floodlight Lux|Integer|1500|
|12|Construction Cost MYR|Currency|RM800,000,000.00|
|13|Has Athletics Track|Text (Yes/No)|Yes|
|14|Last Renovation Date|Date|2017-07-01|
|15|Operator|Text|Perbadanan Stadium Malaysia|
|16|Official Website|Hyperlink|https://www.stadium.gov.my|
|17|Created By|Text|(system field)|
|18|Created On|Date and time|(system field)|

## Workbook structure

|Sheet|Purpose|
|-|-|
|Menu|Home page with buttons to every sheet|
|Form|Data entry form with validation, Save and Reset|
|Database|Excel Table (`DATA`) holding all 50 records|
|Search|Live search across all columns|
|Analysis|KPI cards, PivotTables, PivotCharts and slicers|
|About|Author and project summary with live figures|

## Screenshots

**Menu**

!\[Menu sheet](images/menu.png)

**Data entry form**

!\[Data entry form](images/form.png)

**Database (Excel Table with hyperlinks)**

!\[Database sheet](images/database.png)

**Analysis**

!\[Analysis dashboard](images/analysis.png)

**Search**

!\[Search sheet](images/search.png)

## Key findings

* **Kuala Lumpur has the most venues (9)**, followed by Selangor (7) and Johor (5). These three hold 21 of the 50 venues (42%).
* **68% of venues are outdoor** (34 of 50); indoor arenas make up 24%.
* **Cow Grass is the most common surface** (27 venues, 54%).
* **The 1990s were the busiest decade** with 21 openings, and 1998 alone had 8, matching the Bukit Jalil complex built for the 1998 Commonwealth Games.
* **Venues with an athletics track average about 24,783 seats**, against about 10,964 without one. Most of that gap comes from indoor and aquatic venues, which have no track and are smaller by design, so this is an association and not a cause.
* **Total seating capacity is 948,943**, with an average of 18,979 and a median of 13,000. A few very large stadiums pull the average above the median.

|By state|By venue type|
|-|-|
|!\[Stadiums by state](images/chart-stadiums-by-state.jpg)|!\[Venue type](images/chart-venue-type.png)|

|By surface|By opening decade|
|-|-|
|!\[Surface type](images/chart-surface-type.jpg)|!\[Openings by decade](images/chart-openings-by-decade.jpg)|

!\[Average capacity by athletics track](images/chart-capacity-by-track.jpg)

## How to use the workbook

1. Download `DATABASE\_STADIUM\_IN\_MALAYSIA.xlsm`.
2. Open it in **Microsoft Excel for Windows** (Microsoft 365 or Excel 2021 or later, because the Search sheet uses `FILTER` and the Analysis sheet uses `UNIQUE`).
3. Click **Enable Content** when prompted, so the macros and buttons work.
4. Start from the **Menu** sheet. Use **Register Stadium** to add a record, **Search** to look one up and **Analysis** for the dashboard.

The navigation buttons are ActiveX controls, so they work on Windows only. On Mac, the sheets and data can still be opened and read.

## Limitations

* Excel suits a small, single-user dataset. It does not enforce relational integrity or handle many users at once, so a full database system would be needed at larger scale.
* Some fields are approximate (see the data quality note above).

## Possible next steps

* Move the data into a SQL database (for example PostgreSQL) and write queries for the same questions.
* Rebuild the Analysis sheet as an interactive Power BI dashboard.

## Author

**Nasrizal Fadhli**, Information Systems Management student at UiTM, working towards a data and business analyst career.

