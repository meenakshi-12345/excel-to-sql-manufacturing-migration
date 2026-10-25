# Manufacturing Plant Database: Excel to SQL Server

## A relational database for an industrial machinery components manufacturer, built in Microsoft SQL Server. The company tracked everything in Excel files. I cleaned the raw data, designed the schema, migrated about 10,000 records, added triggers that enforce plant rules, and wrote monthly report views for management.

Tools: SQL Server, T-SQL (CTEs, joins, window-free reporting views, triggers, constraints), SSMS, Excel

## Business Problem

Raw materials, suppliers, production schedules, machines, quality checks and customer orders lived in separate Excel files. This caused:

No unique IDs, so a component could not be traced to its material batch or supplier
Production orders pointing to machines and batches that were not backed by any real record
Inconsistent entries (Kg vs KG, one supplier spelled in many ways)
Machines double-booked for the same time slot
Manual, error-prone reporting

## The full requirement document is in the BRD folder.

Project Structure
├── BRD/                       Business requirement document
├── ABC_10000.xlsx             Original raw data
├── manufacture_database.sql   Full script: tables, cleaning, migration, triggers, reports
└── Schema/                    Screenshots used in this README
Database Design

10 tables connected with primary and foreign keys.

## Database schema

Table	Purpose
suppliers, customers, raw_material	Master lists
machines	Machine name, type, plant, last maintenance date
employees	Inspectors and staff
material_inventory	Each received batch with initial and remaining quantity
production_orders	Orders with schedule, status, customer and assigned machine
production_material_usage	Which batch was used for which order, and how much
quality_check	Inspection result per production order
material_low_stock_log	Alerts created automatically when stock runs low

## I sketched the design on paper first:

Hand-drawn schema

<img width="1280" height="960" alt="handwritten_schema" src="https://github.com/user-attachments/assets/dd52548e-5f87-4fa0-9862-73199baea4eb" />

Digitally Created Schema

<img width="2960" height="1880" alt="schema_diagram" src="https://github.com/user-attachments/assets/1750aed5-ba6b-4ef8-bec1-bce145cf50ee" />


## Data Cleaning

The raw sheet was imported as manufacture_raw_data and left untouched. All cleaning happens once in a staging table (staging_clean), and the final tables are loaded from it.

Issue found	What I did
SupplierID had only 6 distinct values across 10,000+ rows	It could not identify suppliers, so I did not carry it over. The database generates new IDs
Supplier names written many ways (PlasticPro, Copper Co., AluWorks...)	Standardized to 6 clean names with CASE WHEN ... LIKE
CustomerOrderID had 14 values for 98 customer names	Same problem, so customers get a generated ID
One MachineID belonged to several machine names	Dropped it. Plant (P1 / P2) is taken from the machine name prefix
Mixed date formats	Converted with TRY_CONVERT and COALESCE
Quantities stored as text with $, commas and spaces	Cleaned and cast to DECIMAL(10,2)
Units like Kg, KGS, Kilogram, Feet, Square Feet	Standardized to KG, M, PCS, FT, SQ_FT, LBS, enforced by a CHECK constraint
Production IDs written with /, _ and -	Made consistent
Orders with end date before start date, or start date before order date	Left out of the migration and blocked going forward by CHECK constraints



## Business Rules (Triggers)
Rule	How it works
Failed check goes to rework	When a Failed quality result is added, the production order status becomes Rework Required
Stock stays in sync	Adding, changing or deleting material usage updates remaining_quantity in material_inventory. Usage larger than the stock left in the batch is rejected with INSUFFICIENT MATERIAL
No double-booked machines	A machine cannot have two overlapping active orders. Completed and cancelled orders are ignored
Low stock alert	When a batch drops below 500, a row is added to material_low_stock_log with the quantity needed to bring it back to 4,000

Monthly Reports

Three views answer the management questions from the brief.

View	Question it answers
vw_machine_efficiency_monthly	Per machine and month: completion rate, rework rate, scheduled hours
vw_material_consumption_monthly	How much of each material each plant used per month
vw_supplier_quality_scorecard	Which supplier has the lowest defect rate

Machine efficiency here is based on completion and rework rates. A full OEE needs downtime data, which is not recorded in the source.


## How to Run
Create an empty SQL Server database.
Import ABC_10000.xlsx as a table named manufacture_raw_data, and the employee list as employees_raw (SSMS Import Wizard). Use the same table names, the script depends on them.
Run manufacture_database.sql from top to bottom. Each section is separated by GO.
Use the check queries in Part 4 of the script to compare row counts with the raw data.
Still To Do
Role-based access (machine operators, line supervisors, procurement)
Block production orders when the required material is not in stock
Block scheduling on machines under maintenance
About Me

Meenakshi Rajpurohit | B.Tech CSE graduate, aspiring Data Analyst | SQL, Python, Power BI, Excel

