USE UniTrac
BEGIN TRY
    -- Create and populate the temp table
    SELECT TOP 1 * INTO #LFPTestingJLW FROM loan;
    PRINT 'Temp table #LFPTestingJLW created'
END TRY
BEGIN CATCH
    PRINT 'Error while creating temp table #LFPTestingJLW'
END CATCH

-- Your existing SQL script goes here

BEGIN TRY
    -- Drop the temporary table
    DROP TABLE #LFPTestingJLW;
    PRINT 'Dropped temp table #LFPTestingJLW'
END TRY
BEGIN CATCH
    PRINT 'Error while dropping temp table #LFPTestingJLW'
END CATCH
