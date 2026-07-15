/* ============================================================
   Mục tiêu: tách "thời điểm nhập thật" (Time_line) khỏi
   "ngày làm việc / ca" (Work_Date, Shift) bằng computed column
   persisted, để:
     - Time_line không còn bị fake giờ (Controller sẽ luôn lưu
       DateTime.Now thật).
     - Lịch sử / báo cáo lọc & nhóm theo Work_Date, nên bản ghi
       ca đêm nhập sau 00:00 vẫn hiển thị đúng dưới ngày làm việc
       hôm trước, thay vì rơi sang ngày hôm sau theo Time_line.

   Rule ngày/ca (giữ nguyên logic gốc của SP):
     Day   : Time_line giờ trong (08:00, 20:00]
     Night : còn lại; nếu giờ <= 08:00 thì Work_Date lùi 1 ngày

   Chạy 1 lần trên DB svn_pentaho. An toàn với dữ liệu cũ vì
   ALTER TABLE ADD ... PERSISTED tự tính lại cho toàn bộ dòng
   hiện có, không cần backfill riêng.
   ============================================================ */

------------------------------------------------------------
-- 1. Thêm 2 cột computed persisted vào bảng History
------------------------------------------------------------
-- Lưu ý: không so sánh CAST(Time_line AS TIME) với chuỗi '08:00'/'20:00' —
-- implicit convert chuỗi -> TIME phụ thuộc DATEFORMAT/LANGUAGE của session,
-- nên SQL Server coi là non-deterministic và từ chối PERSISTED. Thay vào đó
-- so Time_line (DATETIME) với mốc 08:00/20:00 dựng từ chính ngày của
-- Time_line bằng DATEADD số nguyên giờ — deterministic, và vì cùng ngày nên
-- kết quả so sánh tương đương hệt so theo giờ-trong-ngày như bản gốc.
ALTER TABLE dbo.SVN_Defect_Record_History
ADD Work_Date AS (
        CASE WHEN Time_line <= DATEADD(HOUR, 8, CAST(CAST(Time_line AS DATE) AS DATETIME))
             THEN DATEADD(DAY, -1, CAST(Time_line AS DATE))
             ELSE CAST(Time_line AS DATE)
        END
    ) PERSISTED,
    Shift AS (
        CASE WHEN Time_line >  DATEADD(HOUR, 8,  CAST(CAST(Time_line AS DATE) AS DATETIME))
              AND Time_line <= DATEADD(HOUR, 20, CAST(CAST(Time_line AS DATE) AS DATETIME))
             THEN N'Day' ELSE N'Night'
        END
    ) PERSISTED;
GO

------------------------------------------------------------
-- 2. Index hỗ trợ filter/group theo ngày làm việc + ca
------------------------------------------------------------
CREATE INDEX IX_SVN_Defect_Record_History_WorkDate_Shift
    ON dbo.SVN_Defect_Record_History (Work_Date, Shift);
GO

------------------------------------------------------------
-- 3. SP rút gọn — đọc thẳng Work_Date/Shift đã persisted
--    thay vì tính lại bằng CASE + range filter trong MERGE.
--    Logic tương đương 100% với bản gốc (đã đối chiếu từng
--    điều kiện biên với sp_helptext bản cũ).
------------------------------------------------------------
CREATE OR ALTER PROCEDURE [dbo].[SVN_InsertDefectReport]
    @Work_order     NVARCHAR(50),
    @Item_code      NVARCHAR(50),
    @Defect_Code    NVARCHAR(50),
    @Defect_Name    NVARCHAR(100),
    @Qty_NG         INT,
    @INSDatetime    NVARCHAR(50),
    @Operation      NVARCHAR(50),
    @Employer_code  NVARCHAR(50),
    @Employer_name  NVARCHAR(50),
    @Note           NVARCHAR(MAX),
    @Image_error    NVARCHAR(MAX),
    @Time_line      DATETIME
AS
BEGIN
    SET NOCOUNT ON;

    ------------------ 1. Xác định Shift và ngày làm việc (dùng để lọc ở bước 3) ------------------
    -- Day   : 08:00:01 -> 20:00:00
    -- Night : 20:00:01 -> 08:00:00 hôm sau
    DECLARE @Gio TIME = CAST(@Time_line AS TIME);
    DECLARE @Shift NVARCHAR(10);
    DECLARE @WorkDate DATE = CAST(@Time_line AS DATE);

    IF @Gio > '08:00' AND @Gio <= '20:00'
        SET @Shift = N'Day';
    ELSE
    BEGIN
        SET @Shift = N'Night';
        IF @Gio <= '08:00'  -- 00:00 -> 08:00:00: vẫn thuộc ca đêm của hôm trước
            SET @WorkDate = DATEADD(DAY, -1, @WorkDate);
    END

    ------------------ 2. Insert vào History (Work_Date, Shift tự tính từ Time_line) ------------------
    INSERT INTO dbo.SVN_Defect_Record_History
    (
        [Work_order], [Item_code], [Defect_Code], [Defect_Name],
        [Qty_NG], [INSDatetime], [Operation],
        [Employer_code], [Employer_name], [Note], [Image_error], [Time_line]
    )
    VALUES
    (
        @Work_order, @Item_code, @Defect_Code, @Defect_Name,
        @Qty_NG, @INSDatetime, @Operation,
        @Employer_code, @Employer_name, @Note, @Image_error, @Time_line
    );

    ------------------ 3. MERGE vào Record (đọc thẳng Work_Date/Shift đã persisted) ------------------
    MERGE INTO dbo.SVN_Defect_Record AS target
    USING
    (
        SELECT item_code, defect_code, Operation, Work_Date, Shift,
               SUM(CAST(qty_NG AS DECIMAL)) AS Qty_NG
        FROM dbo.SVN_Defect_Record_History
        WHERE Work_Date = @WorkDate AND Shift = @Shift
        GROUP BY item_code, defect_code, Operation, Work_Date, Shift
    ) AS source
        ON  target.defect_code = source.defect_code
        AND target.item_code   = source.item_code
        AND target.INSDatetime = CONVERT(NVARCHAR(50), source.Work_Date, 23)
        AND target.Operation   = source.Operation
        AND target.Shift       = source.Shift
    WHEN MATCHED THEN
        UPDATE SET target.qty_NG = source.Qty_NG
    WHEN NOT MATCHED BY TARGET THEN
        INSERT (item_code, defect_code, Qty_NG, INSDatetime, Operation, Shift)
        VALUES (source.item_code, source.defect_code, source.Qty_NG,
                CONVERT(NVARCHAR(50), source.Work_Date, 23), source.Operation, source.Shift);
END;
GO
