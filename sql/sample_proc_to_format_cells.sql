
CREATE OR REPLACE PROCEDURE sample_proc_to_format_cells(
p_directory varchar,
p_report_name varchar,
p_LambdaFunctionName varchar,
INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book,
INOUT preturnval varchar DEFAULT NULL
)

LANGUAGE 'plpgsql'
AS $BODY$
/* This procedure is called from PROCP_PM1_REPORT. */
DECLARE
   rec record;
    cell_alignment_center pgexcel_generator.tp_alignment;
    cell_alignment_left pgexcel_generator.tp_alignment;
    cell_alignment_right pgexcel_generator.tp_alignment;
    cell_borderh___ INTEGER;
    /* TopThin, BottomNone, LeftNone, RightNone */
    cell_border_h__ INTEGER;
    /* TopNone, BottomThin, LeftNone, RightNone */
    cell_borderhh__ INTEGER;
    /* TopThin, BottomThin, LeftNone, RightNone */
    cell_bordermm_h INTEGER;
    /* Topmedium, Bottommedium, LeftNone, RightThin */
    cell_bordermm_m INTEGER;
    /* Topmedium, Bottommedium, LeftNone, Rightmedium */
    cell_bordermmm_ INTEGER;
    /* Topmedium, Bottommedium, Leftmedium, RightNone */
    cell_bordermmmm INTEGER;
    /* Topmedium, Bottommedium, Leftmedium, Rightmedium */
    cell_fontAriel8BI INTEGER;
    cell_fontAriel9I INTEGER;
    cell_fontAriel9BI INTEGER;
    cell_fontAriel10I INTEGER;
    cell_fontCalibri9 INTEGER;
    cell_fontMSSS10 INTEGER;
    numFmt_0 INTEGER;
    numFmt_1 INTEGER;
    numFmt_2 INTEGER;
    numFmt_4 INTEGER;
    numFmt_8 INTEGER;
    nAdjustments NUMERIC(18, 2);
    /* This is a placeholder. Adjustments (if any) are added manually once the spreadsheet is available. */
    nATS_Min NUMERIC(18, 2);
    cATSVolRt CHARACTER VARYING(21);
    nDomL3 NUMERIC(18, 2);
    nDomL3Adj NUMERIC(18, 2);
    nDomInter NUMERIC(18, 2);
    nDomRbt NUMERIC(18, 2);
    nDomStd NUMERIC(18, 2);
    nDomWkrb NUMERIC(18, 2);
    nDSO NUMERIC(18, 2);
    nDSOA NUMERIC(18, 2);
    cDSORt CHARACTER VARYING(21);
    nIntInter NUMERIC(18, 2);
    nIntl_Pct NUMERIC(18, 5);
    nIntL3 NUMERIC(18, 2);
    nIntL3Adj NUMERIC(18, 2);
    nIntlRbt NUMERIC(18, 2);
    nIntStd NUMERIC(18, 2);
    nIntWkrb NUMERIC(18, 2);
    cL3Disp CHARACTER VARYING(21);
    nLIII_Pct NUMERIC(18, 8);
    nNbrTxns NUMERIC(18, 0);
    nNDomSpd NUMERIC(18, 2);
    nNIntSpd NUMERIC(18, 2);
    cPyAdj CHARACTER VARYING(24);
    nPyConvRt NUMERIC(18, 8);
    cPyCurr CHARACTER(3);
    nPyRbt NUMERIC(18, 2);
    cQualif CHARACTER(1);
    cQualifN CHARACTER VARYING(106);
    nRbModel NUMERIC(10, 0);
    cRelNm CHARACTER VARYING(100);
    nTATSAmt NUMERIC(16, 2);
    cTierCurr CHARACTER VARYING(3);
    nTtlInter NUMERIC(18, 2);
    nTtlRbt NUMERIC(18, 2);
    cTtlRtDom CHARACTER VARYING(21);
    cTtlRtIntl CHARACTER VARYING(17);
    nTtlSpnd NUMERIC(18, 2);
    nTtlWkrb NUMERIC(18, 2);
    cVndrID CHARACTER VARYING(50);
    cVolRtDom CHARACTER VARYING(21);
    cVolRtInt CHARACTER VARYING(21);
    nRowNumber INTEGER := 16;
    vSQLCode VARCHAR;
    vSQLErrM VARCHAR;
    vLambdaFunctionName         VARCHAR;
    vLambdaFunctionParam        VARCHAR;
    vLambdaFunctionStatusCode   INT;
    v_workbook pgexcel_generator.tp_book;
    v_author pgexcel_generator.tp_authors[];
    pworksheet int;

BEGIN

    v_workbook = p_workbook;
    pworksheet := 1;
    /* Create a new sheet. */
    CALL pgexcel_generator.new_sheet('client_report'::TEXT,v_workbook);
    CALL pgexcel_generator.set_column_width(1, 27.86, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(2, 13.57, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(3, 19.57, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(4, 13.14, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(5, 13.14, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(6, 2.57, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(7, 16.29, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(8, 18.29, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(9, 8, pworksheet,v_workbook);
    CALL pgexcel_generator.set_column_width(10, 7, pworksheet,v_workbook);
    cell_alignment_center := pgexcel_generator.get_alignment('center'::TEXT, 'center'::TEXT, NULL);
    cell_alignment_left := pgexcel_generator.get_alignment('center'::TEXT, 'left'::TEXT, NULL);
    cell_alignment_right := pgexcel_generator.get_alignment('center'::TEXT, 'right'::TEXT, NULL);
    select  * from pgexcel_generator.get_border('thin'::TEXT, 'none'::TEXT, 'none'::TEXT, 'none'::TEXT,v_workbook) into rec;
    cell_borderh___ := rec.ret_val;
    v_workbook = rec.p_workbook;

    select  * from pgexcel_generator.get_border('none'::TEXT, 'thin'::TEXT, 'none'::TEXT, 'none'::TEXT,v_workbook) into rec;
    cell_border_h__ := rec.ret_val ;
    v_workbook = rec.p_workbook;

    select * from pgexcel_generator.get_border('thin'::TEXT, 'thin'::TEXT, 'none'::TEXT, 'none'::TEXT,v_workbook) into rec;
    cell_borderhh__ := rec.ret_val;
    v_workbook = rec.p_workbook;

    select  * from pgexcel_generator.get_border('medium'::TEXT, 'medium'::TEXT, 'medium'::TEXT, 'medium'::TEXT,v_workbook) into rec;
    cell_bordermmmm := rec.ret_val ;
    v_workbook = rec.p_workbook;

    select * from pgexcel_generator.get_border('medium'::TEXT, 'medium'::TEXT, 'none'::TEXT, 'thin'::TEXT,v_workbook) into rec;
    cell_bordermm_h := rec.ret_val;
    v_workbook = rec.p_workbook;

    select * from pgexcel_generator.get_border('medium'::TEXT, 'medium'::TEXT, 'none'::TEXT, 'medium'::TEXT,v_workbook) into rec;
    cell_bordermm_m := rec.ret_val;
    v_workbook = rec.p_workbook;

    select * from pgexcel_generator.get_border('medium'::TEXT, 'medium'::TEXT, 'medium'::TEXT, 'none'::TEXT,v_workbook) into rec;
    cell_bordermmm_ := rec.ret_val;
    v_workbook = rec.p_workbook;

    select * from pgexcel_generator.get_border('medium'::TEXT, 'medium'::TEXT, 'medium'::TEXT, 'medium'::TEXT,v_workbook) into rec;
    cell_bordermmmm := rec.ret_val;
    v_workbook = rec.p_workbook;

    select * from pgexcel_generator.get_font('Ariel'::TEXT, 2, 8, 1, FALSE::BOOLEAN, TRUE::BOOLEAN, TRUE::BOOLEAN, NULL,v_workbook) into rec;
        cell_fontAriel8BI := rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_font('Ariel'::TEXT, 2, 9, 1, FALSE::BOOLEAN, TRUE::BOOLEAN, FALSE::BOOLEAN, NULL,v_workbook) into rec;
        cell_fontAriel9I := rec.ret_val ;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_font('Ariel'::TEXT, 2, 9, 1, FALSE::BOOLEAN, TRUE::BOOLEAN, TRUE::BOOLEAN, NULL,v_workbook) into rec;
        cell_fontAriel9BI := rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_font('Ariel'::TEXT, 2, 10, 1, FALSE::BOOLEAN, TRUE::BOOLEAN, FALSE::BOOLEAN, NULL,v_workbook) into rec;
        cell_fontAriel10I := rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_font('Calibri'::TEXT, 2, 9, 1, FALSE::BOOLEAN, FALSE::BOOLEAN, FALSE::BOOLEAN, NULL,v_workbook) into rec;
        cell_fontCalibri9 := rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_font('MS Sans Serif'::TEXT, 2, 10, 1, FALSE::BOOLEAN, FALSE::BOOLEAN, FALSE::BOOLEAN, NULL,v_workbook) into rec;
        cell_fontMSSS10 := rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_numfmt('#,##0'::TEXT,v_workbook) into rec;

        numFmt_0 := rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_numfmt('#,##0.0'::TEXT,v_workbook) into rec;
        numFmt_1 :=  rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_numfmt('#,##0.00'::TEXT,v_workbook)  into rec;
        numFmt_2 :=  rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from pgexcel_generator.get_numfmt('#,##0.0000'::TEXT,v_workbook)  into rec;
        numFmt_4 := rec.ret_val;
        v_workbook = rec.p_workbook;

        select * from   pgexcel_generator.get_numfmt('#,##0.00000000'::TEXT,v_workbook)  into rec;
        numFmt_8 := rec.ret_val;
        v_workbook = rec.p_workbook;

    CALL pgexcel_generator.cell(3, 2, 'Client Name:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 3, 'Vendor #:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 4, 'Period Begin:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 5, 'Period End:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 6, 'Conditions Met:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 7, 'Condition Notes'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 8, 'Currency:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 9, 'Exchange:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(3, 10, 'Total Amount:'::TEXT, NULL::integer, cell_fontAriel8BI, NULL::integer, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 2, 'Client xyz', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 3, 'Vendor abc', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 4, TO_CHAR(now()-'30 days'::interval, 'MM/DD/YYYY'), NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 5, TO_CHAR(now(), 'MM/DD/YYYY'), NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 6, 'N', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 7, 'All requirements not met', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 8, 'USD', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 9, '1.00', numFmt_8, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 10, '1234.5', numFmt_2, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 2, 'Client xyz', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 3, 'Vendor abc', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 4, TO_CHAR(now()-'30 days'::interval, 'MM/DD/YYYY'), NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 5, TO_CHAR(now(), 'MM/DD/YYYY'), NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 6, 'N', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 7, 'All requirements not met', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 8, 'USD', NULL, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 9, '1.00', numFmt_8, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(5, 10, '1234.5', numFmt_2, cell_fontAriel10I, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 2, 5, 2, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 3, 5, 3, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 4, 5, 4, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 5, 5, 5, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 6, 5, 6, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 7, 5, 7, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 8, 5, 8, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 9, 5, 9, pworksheet,v_workbook);
    CALL pgexcel_generator.mergecells(4, 10, 5, 10, pworksheet,v_workbook);
    /* */
    /* Column A */

        CALL pgexcel_generator.cell(1, 13, CONCAT_WS('', 'Amounts below are listed in ', cTierCurr), NULL, cell_fontAriel9I, NULL, NULL, cell_alignment_left, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(1, 15, 'Spend'::TEXT, NULL, cell_fontAriel9BI, NULL, NULL, cell_alignment_left, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(1, 19, 'Non Level III Spend'::TEXT, NULL, cell_fontAriel9BI, NULL, NULL, cell_alignment_left, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(1, 20, 'Level III Spend'::TEXT, NULL, cell_fontAriel9BI, NULL, NULL, cell_alignment_left, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(1, 21, CONCAT_WS('', '50', '% Level III Adjustment'), NULL, cell_fontAriel9BI, NULL, NULL, cell_alignment_left, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(1, 22, 'Total  Spend'::TEXT, NULL, cell_fontAriel9BI, NULL, NULL, cell_alignment_left, pworksheet,v_workbook);

    /* */
    /* Column B */

        CALL pgexcel_generator.cell(2, 14, 'Amount'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_border_h__, cell_alignment_center, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(2, 15, 8499945.00, numFmt_2, cell_fontAriel9BI, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(2, 16, '', numFmt_2, cell_fontCalibri9, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(2, 18, 'Domestic'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_border_h__, cell_alignment_center, pworksheet,v_workbook);

        CALL pgexcel_generator.cell(2, 19, 8476078.00, numFmt_2, cell_fontAriel9I, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(2, 20, 21691.00, numFmt_2, cell_fontAriel9I, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(2, 21, 10846.00, numFmt_2, cell_fontAriel9I, NULL, cell_border_h__, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(2, 22, 8486924.00, numFmt_2, cell_fontAriel9I, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);


    /* */
    /* Column C */

        CALL pgexcel_generator.cell(3, 14, '#'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_border_h__, cell_alignment_center, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(3, 15, 20906, numFmt_0, cell_fontAriel9BI, NULL, NULL, cell_alignment_center, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(3, 18, 'Foreign'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_border_h__, cell_alignment_center, pworksheet,v_workbook);

        CALL pgexcel_generator.cell(3, 19, 2176.00, numFmt_2, cell_fontAriel9I, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(3, 20, '-', numFmt_2, cell_fontAriel9I, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(3, 21, '-', numFmt_2, cell_fontAriel9I, NULL, cell_border_h__, cell_alignment_right, pworksheet,v_workbook);
        CALL pgexcel_generator.cell(3, 22, 2176.00, numFmt_2, cell_fontAriel9I, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);

    /* */
    /* Column D */
    CALL pgexcel_generator.cell(4, 14, 'Avg Size'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_border_h__, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(4, 15, 406.58, numFmt_2, cell_fontAriel9BI, NULL, NULL, cell_alignment_right, pworksheet,v_workbook);
    /* */

    pworksheet := 2;
    CALL pgexcel_generator.new_sheet('exchange_rates'::TEXT,v_workbook);


    /* Column G */
    CALL pgexcel_generator.cell(7, 15, 'Average Trans Size Requirement'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_bordermmm_, cell_alignment_left, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(7, 16, 'Tier Start Amt'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    /* */
    /* Column H */
    CALL pgexcel_generator.cell(8, 15, 'Average Trans Size Requirement'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_bordermm_m, cell_alignment_left, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(8, 16, 'Tier End Amt'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    /* */
    CALL pgexcel_generator.mergecells(7, 15, 8, 15, pworksheet,v_workbook);
    /* */
    /* Column I */
    CALL pgexcel_generator.cell(9, 15, 100::double precision, numFmt_2, cell_fontAriel9BI, NULL, cell_bordermmm_, cell_alignment_right, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(9, 16, 'Domestic'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_bordermmm_, cell_alignment_center, pworksheet,v_workbook);
    /* */
    /* Column J */
    CALL pgexcel_generator.cell(10, 15, ' '::TEXT, NULL, cell_fontAriel9BI, NULL, cell_bordermm_m, cell_alignment_left, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(10, 16, 'Foreign'::TEXT, NULL, cell_fontAriel9BI, NULL, cell_bordermm_m, cell_alignment_center, pworksheet,v_workbook);
    /* */

    CALL pgexcel_generator.cell(7, 17, 0.00, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(8, 17, 416666.00, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(9, 17, 1.0, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(10, 17, 1.0, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);

    CALL pgexcel_generator.cell(7, 18, 416667.00, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(8, 18, 4170000.00, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(9, 18, 0.0000, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);
    CALL pgexcel_generator.cell(10, 18, 0.0000, numFmt_0, cell_fontMSSS10, NULL, cell_bordermmmm, cell_alignment_center, pworksheet,v_workbook);


    pworksheet := 3;

    CALL pgexcel_generator.new_sheet('all_tables'::TEXT,v_workbook);
    --CALL pgexcel_generator.query2sheet(psql => 'select * from pg_class limit 10'::TEXT, p_sheet => 3, p_workbook => v_workbook);
    --CALL pgexcel_generator.query2sheet(p_sql => 'select * from test_excel'::TEXT, p_sheet => 3, p_workbook => v_workbook);
    CALL pgexcel_generator.query2sheet(p_sql => 'select * from pg_tables'::TEXT, p_sheet => 3, p_workbook => v_workbook);

    CALL pgexcel_generator.save(p_directory, p_report_name,v_workbook,v_author);

          p_report_name := replace(p_report_name, '.xlsx','');

            --vLambdaFunctionName := 'arn:aws:lambda:us-east-1:xxxx:function:CreateExcelReport';
            vLambdaFunctionParam := replace('{"key1": "report_name"}', 'report_name', p_report_name);
            SELECT status_code
            into vLambdaFunctionStatusCode
            FROM aws_lambda.invoke(p_LambdaFunctionName, vLambdaFunctionParam::json);
            raise notice 'Report: %, vLambdaFunctionStatusCode: %', p_report_name, vLambdaFunctionStatusCode ;

           IF vLambdaFunctionStatusCode=200 THEN
               preturnval := NULL;
           ELSE
               RAISE USING hint = -20903, message = preturnval, detail = 'Error in Lambda function execution';
           END IF;



    preturnval := 'Success sample_proc_to_format';
    RETURN;
    EXCEPTION
        WHEN others THEN
            vSQLCode := SQLSTATE;
            vSQLErrM := SQLERRM;
            preturnval := CONCAT_WS('', COALESCE(TRIM(preturnval), ' '), ', SQLCODE = ', vSQLCode, ', SQLERRM = ', TRIM(vSQLErrM));
            RAISE USING hint = -20112, message = preturnval, detail = 'User-defined exception';
END;
$BODY$;
