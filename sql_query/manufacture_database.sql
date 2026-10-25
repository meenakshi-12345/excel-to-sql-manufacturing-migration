/*
=====================================================================
 MANUFACTURING PLANT DATABASE  (Microsoft SQL Server)
 Excel -> relational database: schema, cleaning, migration,
 business-rule triggers and monthly report views.

 BEFORE YOU RUN (use an empty database):
   1. Import ABC_10000.xlsx into SQL Server as a table named
      manufacture_raw_data (SSMS: Tasks > Import Flat File / Import Data).
   2. Import the employee list as a table named employees_raw
      (2 columns: full name, role).
   3. Run this file from top to bottom. Each section is separated by GO.
=====================================================================
*/

IF DB_ID('manufacturing_database') IS NULL
    CREATE DATABASE manufacturing_database;
GO

USE manufacturing_database;
GO


/* =====================================================================
   PART 1: TABLES
   ===================================================================== */

CREATE TABLE suppliers (
    supplier_id    INT IDENTITY(1,1) PRIMARY KEY,
    supplier_name  VARCHAR(50) NOT NULL UNIQUE
);

CREATE TABLE customers (
    customer_id    INT IDENTITY(1,1) PRIMARY KEY,
    customer_name  VARCHAR(100) NOT NULL UNIQUE
);

CREATE TABLE raw_material (
    material_id     INT IDENTITY(1,1) PRIMARY KEY,
    material_name   NVARCHAR(100) NOT NULL,
    material_grade  NVARCHAR(50) NULL,
    CONSTRAINT unique_material UNIQUE (material_name, material_grade)
);

CREATE TABLE machines (
    machine_id             INT IDENTITY(1,1) PRIMARY KEY,
    machine_name           NVARCHAR(50) NOT NULL,
    machine_type           NVARCHAR(50) NULL,
    plant_id               NVARCHAR(10) NOT NULL,
    last_maintenance_date  DATE NULL
);

CREATE TABLE employees (
    employee_id  INT IDENTITY(1,1) PRIMARY KEY,
    full_name    NVARCHAR(100) NOT NULL,
    [role]       NVARCHAR(50) NOT NULL
);

CREATE TABLE production_orders (
    production_id         INT IDENTITY(1,1) PRIMARY KEY,
    production_code       NVARCHAR(20) NOT NULL,
    customer_id           INT NOT NULL,
    component_to_produce  NVARCHAR(255) NOT NULL,
    quantity_to_produce   INT NOT NULL,
    order_date            DATETIME NOT NULL,
    schedule_start_date   DATETIME NOT NULL,
    schedule_end_date     DATETIME NOT NULL,
    [status]              NVARCHAR(50) NOT NULL,
    assigned_machine_id   INT NULL,
    plant_id              NVARCHAR(3) NOT NULL,
    FOREIGN KEY (customer_id) REFERENCES customers(customer_id),
    FOREIGN KEY (assigned_machine_id) REFERENCES machines(machine_id),
    CONSTRAINT check_production_status CHECK ([status] IN
        ('Scheduled', 'In Progress', 'Completed', 'On Hold', 'Cancelled',
         'Rework Required', 'Quoted', 'Rework Complete', 'Pending Materials')),
    CONSTRAINT check_order_dates CHECK (schedule_start_date >= order_date),
    CONSTRAINT check_schedule_sequence CHECK (schedule_end_date >= schedule_start_date)
);

CREATE TABLE material_inventory (
    batch_id            INT IDENTITY(1,1) PRIMARY KEY,
    original_batch_id   NVARCHAR(50) NULL,
    material_id         INT NOT NULL,
    supplier_id         INT NOT NULL,
    received_date       DATE NOT NULL,
    initial_quantity    DECIMAL(10,2) NOT NULL,
    remaining_quantity  DECIMAL(10,2) NOT NULL,
    unit                NVARCHAR(10) NOT NULL,
    FOREIGN KEY (material_id) REFERENCES raw_material(material_id),
    FOREIGN KEY (supplier_id) REFERENCES suppliers(supplier_id),
    CONSTRAINT check_unit CHECK (unit IN ('KG', 'M', 'PCS', 'FT', 'SQ_FT', 'LBS'))
);

CREATE TABLE production_material_usage (
    usage_id             INT IDENTITY(1,1) PRIMARY KEY,
    production_order_id  INT NOT NULL,
    batch_id             INT NOT NULL,
    material_id          INT NOT NULL,
    quantity_used        DECIMAL(10,2) NOT NULL,
    FOREIGN KEY (production_order_id) REFERENCES production_orders(production_id),
    FOREIGN KEY (batch_id) REFERENCES material_inventory(batch_id),
    FOREIGN KEY (material_id) REFERENCES raw_material(material_id)
);

CREATE TABLE quality_check (
    quality_check_id           INT IDENTITY(1,1) PRIMARY KEY,
    original_quality_check_id  NVARCHAR(50) NULL,
    production_order_id        INT NOT NULL,
    inspector_id               INT NOT NULL,
    check_timestamp            DATETIME NULL,
    result                     NVARCHAR(10) NULL,
    FOREIGN KEY (production_order_id) REFERENCES production_orders(production_id),
    FOREIGN KEY (inspector_id) REFERENCES employees(employee_id)
);

CREATE TABLE material_low_stock_log (
    logid                   INT IDENTITY(1,1) PRIMARY KEY,
    batch_id                INT NULL,
    material_id             INT NOT NULL,
    material_name           NVARCHAR(100) NULL,
    material_grade          NVARCHAR(50) NULL,
    alert_trigger_quantity  DECIMAL(10,2) NULL,
    quantity_to_order       DECIMAL(10,2) NULL,
    alert_date              DATETIME DEFAULT GETDATE()
);
GO


/* =====================================================================
   PART 2: CLEAN THE RAW DATA INTO A STAGING TABLE

   Findings from the raw data:
   - SupplierID has only 6 distinct values for 10k+ rows, CustomerOrderID has
     14 values for 98 customers, and one MachineID maps to several machine
     names. These IDs cannot identify anything, so they were not carried over.
     New surrogate IDs are generated by the database instead.
   - Supplier names, units, dates and quantities are written in many formats.
   All fixes are done once here, so the inserts below stay simple.
   The original manufacture_raw_data table is left untouched.
   ===================================================================== */

SELECT
    -- supplier
    LTRIM(RTRIM(
        CASE WHEN SupplierName LIKE '%PlasticPro%'   THEN 'Plastic Pro Inc.'
             WHEN SupplierName LIKE '%Steelcorp%'    THEN 'Steel Corp'
             WHEN SupplierName LIKE '%CopperCo%'     THEN 'Copper Company'
             WHEN SupplierName LIKE '%Copper Co.%'   THEN 'Copper Company'
             WHEN SupplierName LIKE '%AluWorks%'     THEN 'Alu Works'
             WHEN SupplierName LIKE '%GlobalMetals%' THEN 'Global Metals'
             ELSE SupplierName
        END)) AS supplier_name,

    -- customer
    CustomerName AS customer_name,

    -- material
    MaterialName  AS material_name,
    MaterialGrade AS material_grade,

    -- machine (plant comes from the P1-/P2- prefix of the machine name)
    TRIM(REPLACE(REPLACE(MachineName, 'P1-', ''), 'P2-', '')) AS machine_name,
    TRIM(MachineType)                                          AS machine_type,
    LEFT(MachineName, 2)                                       AS plant_code,
    TRY_CONVERT(DATE, LastMaintenanceDate)                     AS last_maintenance_date,

    -- inventory batch
    LEFT(RawMaterialBatchID, 7)                                AS raw_batch_id,
    TRY_CONVERT(DATE, ReceiveDate, 105)                        AS receive_date,
    CAST(CAST(REPLACE(REPLACE(REPLACE(InitialQuantity, '$', ''), ',', ''), ' ', '') AS FLOAT) AS DECIMAL(10,2)) AS initial_quantity,
    CASE WHEN UPPER(TRIM(Unit)) IN ('KG', 'KILOGRAM', 'KGS')   THEN 'KG'
         WHEN UPPER(TRIM(Unit)) IN ('M', 'METERS', 'METER')    THEN 'M'
         WHEN UPPER(TRIM(Unit)) IN ('PCS', 'PIECES')           THEN 'PCS'
         WHEN UPPER(TRIM(Unit)) IN ('FT', 'FEET')              THEN 'FT'
         WHEN UPPER(TRIM(Unit)) IN ('SQ FT', 'SQUARE FEET')    THEN 'SQ_FT'
         WHEN UPPER(TRIM(Unit)) = 'LBS'                        THEN 'LBS'
         ELSE NULL
    END AS unit_std,

    -- production order (IDs written with / _ - are made consistent)
    RIGHT(REPLACE(REPLACE(ProductionOrderID, '/', '-'), '_', '-'), 4) AS production_code,
    ComponentToProduce                                                 AS component_to_produce,
    CAST(CEILING(QuantityToProduce) AS INT)                            AS quantity_to_produce,
    COALESCE(TRY_CONVERT(DATETIME, OrderDate, 105),
             TRY_CONVERT(DATETIME, OrderDate, 110))                    AS order_date,
    TRY_CONVERT(DATETIME, ScheduledStart, 105)                         AS schedule_start,
    TRY_CONVERT(DATETIME, ScheduledEnd, 105)                           AS schedule_end,
    CASE WHEN TRIM([Status]) = 'Complete'         THEN 'Completed'
         WHEN TRIM([Status]) = 'Rework Completed' THEN 'Rework Complete'
         ELSE TRIM([Status])
    END AS status_std,

    -- material usage
    MaterialUsedBatchID                 AS material_used_batch_id,
    TRY_CAST(QuantityUsed AS DECIMAL(10,2)) AS quantity_used,

    -- quality check
    QualityCheckID                      AS quality_check_id,
    InspectorName                       AS inspector_name,
    TRY_CONVERT(DATETIME, CheckTimestamp, 105) AS check_timestamp,
    Result                              AS result
INTO staging_clean
FROM manufacture_raw_data;
GO


/* =====================================================================
   PART 3: MIGRATE DATA INTO THE TABLES (parents first)
   ===================================================================== */

-- Suppliers
INSERT INTO suppliers (supplier_name)
SELECT DISTINCT supplier_name
FROM staging_clean
WHERE supplier_name IS NOT NULL;

-- Customers
INSERT INTO customers (customer_name)
SELECT DISTINCT customer_name
FROM staging_clean
WHERE customer_name IS NOT NULL;

-- Raw material (one row per material + grade)
INSERT INTO raw_material (material_name, material_grade)
SELECT DISTINCT material_name, material_grade
FROM staging_clean
WHERE material_name IS NOT NULL;

-- Machines (one row per machine name + type + plant, latest maintenance date)
INSERT INTO machines (machine_name, machine_type, plant_id, last_maintenance_date)
SELECT machine_name, machine_type, plant_code, MAX(last_maintenance_date)
FROM staging_clean
WHERE machine_name IS NOT NULL
GROUP BY machine_name, machine_type, plant_code;

-- Employees (employees_raw: full name, role)
INSERT INTO employees (full_name, [role])
SELECT * FROM employees_raw;

-- Material inventory (remaining quantity starts equal to initial quantity,
-- it is reduced below once usage is loaded)
INSERT INTO material_inventory
    (original_batch_id, material_id, supplier_id, received_date, initial_quantity, remaining_quantity, unit)
SELECT s.raw_batch_id, r.material_id, sp.supplier_id, s.receive_date,
       s.initial_quantity, s.initial_quantity, s.unit_std
FROM staging_clean AS s
INNER JOIN suppliers    AS sp ON sp.supplier_name = s.supplier_name
INNER JOIN raw_material AS r  ON r.material_name = s.material_name
                             AND (r.material_grade = s.material_grade
                                  OR (r.material_grade IS NULL AND s.material_grade IS NULL))
WHERE s.unit_std IS NOT NULL
  AND s.receive_date IS NOT NULL
  AND s.initial_quantity IS NOT NULL;

-- Production orders
-- Rows with impossible dates (end before start, or start before order date)
-- are left out. The CHECK constraints would reject them anyway.
INSERT INTO production_orders
    (production_code, customer_id, component_to_produce, quantity_to_produce,
     order_date, schedule_start_date, schedule_end_date, [status], assigned_machine_id, plant_id)
SELECT s.production_code, c.customer_id, s.component_to_produce, s.quantity_to_produce,
       s.order_date, s.schedule_start, s.schedule_end, s.status_std, m.machine_id, m.plant_id
FROM staging_clean AS s
INNER JOIN customers AS c ON c.customer_name = s.customer_name
INNER JOIN machines  AS m ON m.machine_name = s.machine_name
                         AND ISNULL(m.machine_type, '') = ISNULL(s.machine_type, '')
                         AND m.plant_id = s.plant_code
WHERE s.schedule_end   >= s.schedule_start
  AND s.schedule_start >= s.order_date;

-- Material usage (each staging row is matched to its production order
-- and to its inventory batch)
INSERT INTO production_material_usage (production_order_id, batch_id, material_id, quantity_used)
SELECT p.production_id, mi.batch_id, mi.material_id, s.quantity_used
FROM staging_clean AS s
INNER JOIN production_orders AS p
        ON  p.production_code      = s.production_code
        AND p.schedule_start_date  = s.schedule_start
        AND p.schedule_end_date    = s.schedule_end
        AND p.quantity_to_produce  = s.quantity_to_produce
        AND p.component_to_produce = s.component_to_produce
INNER JOIN material_inventory AS mi
        ON mi.original_batch_id = LEFT(s.material_used_batch_id, 7)
WHERE s.quantity_used IS NOT NULL;

-- Remaining quantity = initial quantity - total used from that batch
UPDATE mi
SET mi.remaining_quantity = mi.initial_quantity - u.total_used
FROM material_inventory AS mi
INNER JOIN (
    SELECT batch_id, SUM(quantity_used) AS total_used
    FROM production_material_usage
    GROUP BY batch_id
) AS u ON u.batch_id = mi.batch_id;

-- Quality checks
-- (production_order_id uses the production_orders primary key, not the 4-digit code)
INSERT INTO quality_check
    (original_quality_check_id, production_order_id, inspector_id, check_timestamp, result)
SELECT s.quality_check_id, p.production_id, e.employee_id, s.check_timestamp, s.result
FROM staging_clean AS s
INNER JOIN production_orders AS p
        ON  p.production_code      = s.production_code
        AND p.schedule_start_date  = s.schedule_start
        AND p.schedule_end_date    = s.schedule_end
        AND p.quantity_to_produce  = s.quantity_to_produce
        AND p.component_to_produce = s.component_to_produce
INNER JOIN employees AS e ON e.full_name = s.inspector_name
WHERE s.quality_check_id IS NOT NULL;

-- Client request: every failed quality check puts the order in 'Rework Required'
UPDATE production_orders
SET [status] = 'Rework Required'
WHERE production_id IN (SELECT production_order_id FROM quality_check WHERE result = 'Failed');
GO


/* =====================================================================
   PART 4: CHECK THE MIGRATION (run these and compare with the raw data)
   ===================================================================== */

SELECT 'manufacture_raw_data' AS table_name, COUNT(*) AS row_count FROM manufacture_raw_data
UNION ALL SELECT 'suppliers',                 COUNT(*) FROM suppliers
UNION ALL SELECT 'customers',                 COUNT(*) FROM customers
UNION ALL SELECT 'raw_material',              COUNT(*) FROM raw_material
UNION ALL SELECT 'machines',                  COUNT(*) FROM machines
UNION ALL SELECT 'employees',                 COUNT(*) FROM employees
UNION ALL SELECT 'material_inventory',        COUNT(*) FROM material_inventory
UNION ALL SELECT 'production_orders',         COUNT(*) FROM production_orders
UNION ALL SELECT 'production_material_usage', COUNT(*) FROM production_material_usage
UNION ALL SELECT 'quality_check',             COUNT(*) FROM quality_check;

-- Batches that appear in more than one inventory row (should be reviewed)
SELECT original_batch_id, COUNT(*) AS rows_for_batch
FROM material_inventory
GROUP BY original_batch_id
HAVING COUNT(*) > 1;

-- Batches where usage is higher than the quantity received
SELECT * FROM material_inventory WHERE remaining_quantity < 0;

-- Machines already double-booked in the historical data (before the trigger existed)
SELECT a.production_id AS order_a, b.production_id AS order_b, a.assigned_machine_id
FROM production_orders AS a
INNER JOIN production_orders AS b
        ON  a.assigned_machine_id = b.assigned_machine_id
        AND a.production_id < b.production_id
        AND a.schedule_start_date < b.schedule_end_date
        AND a.schedule_end_date   > b.schedule_start_date
WHERE a.[status] NOT IN ('Completed', 'Cancelled', 'Rework Complete')
  AND b.[status] NOT IN ('Completed', 'Cancelled', 'Rework Complete');
GO


/* =====================================================================
   PART 5: BUSINESS RULES (TRIGGERS)
   Created after the migration so the historical data loads freely.
   ===================================================================== */

-- Rule 1: a failed quality check sends the production order to rework
CREATE TRIGGER trg_rework_on_failed_qc
ON quality_check
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    UPDATE p
    SET p.[status] = 'Rework Required'
    FROM production_orders AS p
    INNER JOIN inserted AS i ON i.production_order_id = p.production_id
    WHERE i.result = 'Failed'
      AND p.[status] <> 'Rework Required';
END;
GO


-- Rule 2: keep remaining stock in sync with material usage,
--         and block usage that is more than what is left in the batch.
--         Works for INSERT, UPDATE and DELETE, including many rows at once.
CREATE TRIGGER trg_material_usage_stock
ON production_material_usage
AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;

    -- give back the old quantity (rows in "deleted": updated or deleted rows)
    UPDATE mi
    SET mi.remaining_quantity = mi.remaining_quantity + d.total_qty
    FROM material_inventory AS mi
    INNER JOIN (SELECT batch_id, SUM(quantity_used) AS total_qty
                FROM deleted GROUP BY batch_id) AS d
            ON d.batch_id = mi.batch_id;

    -- take off the new quantity (rows in "inserted": new or updated rows)
    UPDATE mi
    SET mi.remaining_quantity = mi.remaining_quantity - i.total_qty
    FROM material_inventory AS mi
    INNER JOIN (SELECT batch_id, SUM(quantity_used) AS total_qty
                FROM inserted GROUP BY batch_id) AS i
            ON i.batch_id = mi.batch_id;

    -- stop if any affected batch went below zero
    IF EXISTS (
        SELECT 1
        FROM material_inventory AS mi
        INNER JOIN (SELECT DISTINCT batch_id FROM inserted) AS i
                ON i.batch_id = mi.batch_id
        WHERE mi.remaining_quantity < 0
    )
    BEGIN
        RAISERROR('INSUFFICIENT MATERIAL', 16, 1);
        ROLLBACK TRANSACTION;
        RETURN;
    END
END;
GO


-- Rule 3: a machine cannot have two overlapping production orders.
--         Finished and cancelled orders are ignored. Only checked when the
--         machine or the schedule changes, so a status update (for example
--         'Rework Required') is never blocked by old overlapping data.
CREATE TRIGGER trg_machine_schedule_conflict
ON production_orders
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT (UPDATE(assigned_machine_id) OR UPDATE(schedule_start_date) OR UPDATE(schedule_end_date))
        RETURN;

    IF EXISTS (
        SELECT 1
        FROM inserted AS i
        INNER JOIN production_orders AS p
                ON  p.assigned_machine_id = i.assigned_machine_id
                AND p.production_id <> i.production_id
        WHERE i.[status] NOT IN ('Completed', 'Cancelled', 'Rework Complete')
          AND p.[status] NOT IN ('Completed', 'Cancelled', 'Rework Complete')
          AND i.schedule_start_date < p.schedule_end_date
          AND i.schedule_end_date   > p.schedule_start_date
    )
    BEGIN
        RAISERROR('Schedule conflict: the assigned machine is already booked in this time period', 16, 1);
        ROLLBACK TRANSACTION;
        RETURN;
    END
END;
GO


-- Rule 4: low stock alert. When a batch drops below 500 for the first time,
--         log it with the quantity needed to bring it back to 4000.
CREATE TRIGGER trg_low_stock
ON material_inventory
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO material_low_stock_log
        (batch_id, material_id, material_name, material_grade, alert_trigger_quantity, quantity_to_order)
    SELECT i.batch_id, i.material_id, r.material_name, r.material_grade,
           i.remaining_quantity, 4000 - i.remaining_quantity
    FROM inserted AS i
    INNER JOIN raw_material AS r ON r.material_id = i.material_id
    LEFT JOIN deleted AS d ON d.batch_id = i.batch_id
    WHERE i.remaining_quantity < 500
      AND (d.batch_id IS NULL OR d.remaining_quantity >= 500);
END;
GO


/* =====================================================================
   PART 6: MONTHLY REPORT VIEWS
   ===================================================================== */

-- Report 1: machine efficiency per month
-- (completion rate, rework rate and scheduled hours; this is a simple
--  efficiency measure, not a full OEE, because downtime is not recorded)
CREATE VIEW vw_machine_efficiency_monthly AS
SELECT
    DATEFROMPARTS(YEAR(p.schedule_start_date), MONTH(p.schedule_start_date), 1) AS report_month,
    m.plant_id,
    m.machine_id,
    m.machine_name,
    m.machine_type,
    COUNT(*) AS total_orders,
    SUM(CASE WHEN p.[status] IN ('Completed', 'Rework Complete') THEN 1 ELSE 0 END) AS completed_orders,
    SUM(CASE WHEN p.[status] = 'Rework Required' THEN 1 ELSE 0 END) AS rework_orders,
    CAST(100.0 * SUM(CASE WHEN p.[status] IN ('Completed', 'Rework Complete') THEN 1 ELSE 0 END)
         / COUNT(*) AS DECIMAL(5,2)) AS completion_rate_pct,
    CAST(100.0 * SUM(CASE WHEN p.[status] = 'Rework Required' THEN 1 ELSE 0 END)
         / COUNT(*) AS DECIMAL(5,2)) AS rework_rate_pct,
    SUM(DATEDIFF(HOUR, p.schedule_start_date, p.schedule_end_date)) AS scheduled_hours
FROM production_orders AS p
INNER JOIN machines AS m ON m.machine_id = p.assigned_machine_id
GROUP BY DATEFROMPARTS(YEAR(p.schedule_start_date), MONTH(p.schedule_start_date), 1),
         m.plant_id, m.machine_id, m.machine_name, m.machine_type;
GO


-- Report 2: raw material consumption by plant (production line) per month
CREATE VIEW vw_material_consumption_monthly AS
SELECT
    DATEFROMPARTS(YEAR(p.schedule_start_date), MONTH(p.schedule_start_date), 1) AS report_month,
    p.plant_id,
    r.material_name,
    r.material_grade,
    mi.unit,
    SUM(u.quantity_used) AS total_quantity_used,
    COUNT(DISTINCT p.production_id) AS orders_using_material
FROM production_material_usage AS u
INNER JOIN production_orders  AS p  ON p.production_id = u.production_order_id
INNER JOIN material_inventory AS mi ON mi.batch_id = u.batch_id
INNER JOIN raw_material       AS r  ON r.material_id = mi.material_id
GROUP BY DATEFROMPARTS(YEAR(p.schedule_start_date), MONTH(p.schedule_start_date), 1),
         p.plant_id, r.material_name, r.material_grade, mi.unit;
GO


-- Report 3: supplier quality scorecard
-- (a check is counted against a supplier when the checked order used
--  material supplied by that supplier)
CREATE VIEW vw_supplier_quality_scorecard AS
SELECT
    sp.supplier_id,
    sp.supplier_name,
    COUNT(DISTINCT qc.quality_check_id) AS total_checks,
    COUNT(DISTINCT CASE WHEN qc.result = 'Failed' THEN qc.quality_check_id END) AS failed_checks,
    CAST(100.0 * COUNT(DISTINCT CASE WHEN qc.result = 'Failed' THEN qc.quality_check_id END)
         / NULLIF(COUNT(DISTINCT qc.quality_check_id), 0) AS DECIMAL(5,2)) AS defect_rate_pct
FROM quality_check AS qc
INNER JOIN production_material_usage AS u  ON u.production_order_id = qc.production_order_id
INNER JOIN material_inventory        AS mi ON mi.batch_id = u.batch_id
INNER JOIN suppliers                 AS sp ON sp.supplier_id = mi.supplier_id
GROUP BY sp.supplier_id, sp.supplier_name;
GO


/* =====================================================================
   PART 7: EXAMPLE QUERIES AND TRIGGER TESTS
   (run one at a time; the batch / order numbers must exist in your data)
   ===================================================================== */

-- Reports
-- SELECT * FROM vw_machine_efficiency_monthly ORDER BY report_month, plant_id, rework_rate_pct DESC;
-- SELECT * FROM vw_material_consumption_monthly ORDER BY report_month, plant_id, total_quantity_used DESC;
-- SELECT * FROM vw_supplier_quality_scorecard ORDER BY defect_rate_pct DESC;

-- Test rule 1: failed check -> order goes to 'Rework Required'
-- INSERT INTO quality_check (original_quality_check_id, production_order_id, inspector_id, check_timestamp, result)
-- VALUES ('QC-TEST', 1, 1, GETDATE(), 'Failed');
-- SELECT production_id, [status] FROM production_orders WHERE production_id = 1;

-- Test rule 2: stock goes down, and too much usage is rejected
-- SELECT batch_id, material_id, remaining_quantity FROM material_inventory WHERE batch_id = 105;
-- INSERT INTO production_material_usage (production_order_id, batch_id, material_id, quantity_used)
-- VALUES (5, 105, 573, 100);       -- 573 = material_id of batch 105 in your data; stock drops by 100
-- INSERT INTO production_material_usage (production_order_id, batch_id, material_id, quantity_used)
-- VALUES (5, 105, 573, 99999999);  -- more than available: error INSUFFICIENT MATERIAL

-- Test rule 3: pick an active order and try to book another order on the same
-- machine in the same time slot; the trigger should reject it.

-- Test rule 4: SELECT * FROM material_low_stock_log;
