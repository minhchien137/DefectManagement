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
-- 3. SP — fix so với bản trước:
--    a) target.INSDatetime của SVN_Defect_Record LUÔN là format
--       yyyyMMdd (không dấu gạch ngang) — y hệt bản gốc trước khi
--       có ca ngày/đêm (SP cũ insert thẳng INSdatetime dạng này,
--       và toàn bộ dòng cũ trong bảng đang ở format này). Bản
--       trước dùng CONVERT(...,23) ra "yyyy-MM-dd" nên KHÔNG BAO
--       GIỜ khớp với dòng cũ, và tạo dòng mới format lệch — đây
--       là nguyên nhân Record không nhận được dữ liệu đúng.
--    b) Không phụ thuộc cột computed Work_Date/Shift persisted
--       trên History nữa — tính thẳng từ Time_line ngay trong
--       subquery của MERGE, để chắc chắn không lỗi dù cột đó có
--       tồn tại hay không.
--    c) Bọc transaction (SET XACT_ABORT ON + BEGIN/COMMIT TRAN):
--       nếu bước MERGE lỗi thì insert History cũng rollback theo,
--       không còn tình trạng History có data mà Record thì không.
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
    SET XACT_ABORT ON;

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

    -- format yyyyMMdd, đồng bộ với INSDatetime cũ trong SVN_Defect_Record
    DECLARE @WorkDateStr NVARCHAR(8) = CONVERT(NVARCHAR(8), @WorkDate, 112);

    BEGIN TRANSACTION;

    ------------------ 2. Insert vào History ------------------
    INSERT INTO dbo.SVN_Defect_Record_History
    (
        [Work_order], [Item_code], [Defect_Code], [Defect_Name],
        [Qty_NG], [INSDatetime], [Operation],
        [Employer_code], [Employer_name], [Note], [Image_error], [Time_line]
    )
    VALUES
    (
        @Work_order, @Item_code, @Defect_Code, @Defect_Name,
        @Qty_NG, @WorkDateStr, @Operation,
        @Employer_code, @Employer_name, @Note, @Image_error, @Time_line
    );

    ------------------ 3. MERGE vào Record — tự tính Work_Date/Shift từ Time_line ------------------
    MERGE INTO dbo.SVN_Defect_Record AS target
    USING
    (
        SELECT item_code, defect_code, Operation,
               SUM(CAST(qty_NG AS DECIMAL)) AS Qty_NG
        FROM dbo.SVN_Defect_Record_History
        WHERE
            (CASE WHEN Time_line <= DATEADD(HOUR, 8, CAST(CAST(Time_line AS DATE) AS DATETIME))
                  THEN DATEADD(DAY, -1, CAST(Time_line AS DATE))
                  ELSE CAST(Time_line AS DATE)
             END) = @WorkDate
            AND
            (CASE WHEN Time_line >  DATEADD(HOUR, 8,  CAST(CAST(Time_line AS DATE) AS DATETIME))
                   AND Time_line <= DATEADD(HOUR, 20, CAST(CAST(Time_line AS DATE) AS DATETIME))
                  THEN N'Day' ELSE N'Night'
             END) = @Shift
        GROUP BY item_code, defect_code, Operation
    ) AS source
        ON  target.defect_code = source.defect_code
        AND target.item_code   = source.item_code
        AND target.INSDatetime = @WorkDateStr
        AND target.Operation   = source.Operation
        AND target.Shift       = @Shift
    WHEN MATCHED THEN
        UPDATE SET target.qty_NG = source.Qty_NG
    WHEN NOT MATCHED BY TARGET THEN
        INSERT (item_code, defect_code, Qty_NG, INSDatetime, Operation, Shift)
        VALUES (source.item_code, source.defect_code, source.Qty_NG,
                @WorkDateStr, source.Operation, @Shift);

    COMMIT TRANSACTION;
END;
GO
