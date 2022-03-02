
--
-- Name: xlsx_builder_pkg; Type: SCHEMA; Schema: -; Owner: -
--

CREATE SCHEMA pgexcel_generator;


--
-- Name: tp_alignment; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_alignment AS (
	vertical character varying,
	horizontal character varying,
	wraptext boolean
);

--
-- Name: tp_authors; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_authors AS (
	col_ind character varying,
	column_value integer
);


--
-- Name: tp_autofilter; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_autofilter AS (
	column_start integer,
	column_end integer,
	row_start integer,
	row_end integer
);


--
-- Name: tp_border; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_border AS (
	top character varying,
	bottom character varying,
	"left" character varying,
	"right" character varying
);


--
-- Name: tp_cell; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_cell AS (
	value double precision,
	style character varying
);


--
-- Name: tp_cells; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_cells AS (
	cell pgexcel_generator.tp_cell[]
);


--
-- Name: tp_comment; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_comment AS (
	text character varying,
	author character varying,
	"row" integer,
	"column" integer,
	width integer,
	height integer
);


--
-- Name: tp_defined_name; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_defined_name AS (
	name character varying,
	ref character varying,
	sheet integer
);


--
-- Name: tp_fill; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_fill AS (
	patterntype character varying,
	fgrgb character varying
);


--
-- Name: tp_font; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_font AS (
	name character varying,
	family integer,
	fontsize double precision,
	theme integer,
	rgb character varying,
	underline boolean,
	italic boolean,
	bold boolean
);


--
-- Name: tp_hyperlink; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_hyperlink AS (
	cell character varying,
	url character varying
);


--
-- Name: tp_numfmt; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_numfmt AS (
	numfmtid integer,
	formatcode character varying
);


--
-- Name: tp_validation; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_validation AS (
	type character varying,
	errorstyle character varying,
	showinputmessage boolean,
	prompt character varying,
	title character varying,
	error_title character varying,
	error_txt character varying,
	showerrormessage boolean,
	formula1 character varying,
	formula2 character varying,
	allowblank boolean,
	sqref character varying
);


--
-- Name: tp_xf_fmt; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_xf_fmt AS (
	numfmtid integer,
	fontid integer,
	fillid integer,
	borderid integer,
	alignment pgexcel_generator.tp_alignment
);


--
-- Name: tp_sheet; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_sheet AS (
	rows pgexcel_generator.tp_cells[],
	widths numeric[],
	name character varying(100),
	freeze_rows integer,
	freeze_cols integer,
	autofilters pgexcel_generator.tp_autofilter[],
	hyperlinks pgexcel_generator.tp_hyperlink[],
	col_fmts pgexcel_generator.tp_xf_fmt[],
	row_fmts pgexcel_generator.tp_xf_fmt[],
	comments pgexcel_generator.tp_comment[],
	mergecells character varying[],
	validations pgexcel_generator.tp_validation[]
);


--
-- Name: tp_string; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_string AS (
	col_ind text,
	col_val numeric
);


--
-- Name: tp_book; Type: TYPE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE TYPE pgexcel_generator.tp_book AS (
	sheets pgexcel_generator.tp_sheet[],
	strings pgexcel_generator.tp_string[],
	str_ind text[],
	str_cnt integer,
	fonts pgexcel_generator.tp_font[],
	fills pgexcel_generator.tp_fill[],
	borders pgexcel_generator.tp_border[],
	numfmts pgexcel_generator.tp_numfmt[],
	cellxfs pgexcel_generator.tp_xf_fmt[],
	numfmtindexes numeric[],
	defined_names pgexcel_generator.tp_defined_name[]
);

--
-- Name: instr(text, text,numeric, numeric); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.instr(p_str text, p_substr text, p_pos numeric DEFAULT 1, p_occurrence numeric DEFAULT 1)
 RETURNS numeric
 LANGUAGE sql
 IMMUTABLE PARALLEL SAFE STRICT
AS $function$


    SELECT
        CASE

            WHEN LENGTH($1) = 0 OR LENGTH($2) = 0 THEN NULL::NUMERIC

            WHEN TRUNC($4) = 0 THEN 1/TRUNC($4)

            WHEN $4 < 0 THEN SQRT($4)

            ELSE

                (

                    WITH RECURSIVE t(str, shift, pos, tail, o, n) AS
                    (
                        SELECT CASE WHEN TRUNC($3) < 0 THEN REVERSE($1) ELSE $1 END AS str,
                            0 AS shift,
                            CASE WHEN TRUNC($3) < 0 THEN -1 * TRUNC($3)::INT ELSE TRUNC($3)::INT END AS pos,
                            CASE WHEN TRUNC($3) < 0 THEN REVERSE($1) ELSE $1 END AS tail,
                            0 AS o,
                            CASE WHEN TRUNC($3) < 0 THEN REVERSE($2) ELSE $2 END AS n
                        UNION ALL
                        SELECT str,
                            shift + pos AS shift,
                            STRPOS(SUBSTR(str, shift + pos), n) AS pos,
                            SUBSTR(str, shift + pos) AS tail,
                            o + 1 AS o,
                            n
                        FROM t
                        WHERE pos <> 0
                    )
                    ,r AS
                    (
                        SELECT t.str,
                            t.shift,
                            t.pos,
                            t.tail,
                            t.o,
                            CASE
                                WHEN TRUNC($3) > 0 THEN
                                    t.pos + t.shift - 1
                                ELSE
                                    LENGTH(t.str) - t.pos - t.shift + 2
                            END cc
                        FROM t
                        WHERE t.o = TRUNC($4)
                        AND t.pos <> 0
                    )
                    SELECT COALESCE
                    (
                        (
                            SELECT r.cc
                            FROM r
                        ),
                        0
                    )::NUMERiC
                )

        END;

$function$;

--
-- Name: substr(text, numeric, numeric); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.substr(text, numeric, numeric)
 RETURNS text
 LANGUAGE sql
 IMMUTABLE STRICT
AS $function$

    SELECT
        CASE
            WHEN TRUNC($3)::INTEGER <= 0 THEN

                NULL::TEXT

            WHEN ABS(TRUNC($2)::INTEGER) > LENGTH($1) THEN

                NULL::TEXT

            WHEN TRUNC($2)::INTEGER >= 0 THEN

                SUBSTR($1, CASE WHEN TRUNC($2)::INTEGER = 0 THEN 1 ELSE TRUNC($2)::INTEGER END, TRUNC($3)::INTEGER)

            ELSE

                SUBSTR($1, LENGTH($1) + TRUNC($2)::INTEGER + 1, TRUNC($3)::INTEGER)
        END;

$function$;

--
-- Name: substr(text, numeric); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--


CREATE OR REPLACE FUNCTION pgexcel_generator.substr(text, numeric)
 RETURNS text
 LANGUAGE sql
 IMMUTABLE STRICT
AS $function$

    SELECT
        CASE
            WHEN ABS(TRUNC($2)::INTEGER) > LENGTH($1) THEN

                NULL::TEXT

            WHEN TRUNC($2)::INTEGER >= 0 THEN

                SUBSTR($1, CASE WHEN TRUNC($2)::INTEGER = 0 THEN 1 ELSE TRUNC($2)::INTEGER END)

            ELSE

                SUBSTR($1, LENGTH($1) + TRUNC($2)::INTEGER + 1)
        END;

$function$;

--
-- Name: add_string(text, pgexcel_generator.tp_book); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.add_string(p_string text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, OUT ret_val integer) RETURNS record
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;    t_cnt INTEGER;
    v_workbook pgexcel_generator.tp_book = p_workbook;
	-- r record;
BEGIN
    --PERFORM pgexcel_generator.Init();
    IF (select  count(*) from unnest (v_workbook.strings) where col_ind = p_string) > 0   THEN
        BEGIN
           select col_Val from  unnest (v_workbook.strings) where col_ind = p_string into t_cnt;
        END;
    ELSE
        t_cnt := COALESCE(array_length(v_workbook.strings, 1), 0);
        v_workbook.str_ind[t_cnt] := p_string;
        --if COALESCE(array_length(v_workbook.strings, 1), 0)  = 0
		  if t_cnt = 0
			then
            v_workbook.strings[1] = row(p_string, t_cnt);
        else
              v_workbook.strings[t_cnt+1] = row(p_string, t_cnt);
        end if;
    END IF;
    v_workbook.str_cnt := v_workbook.str_cnt + 1;
    p_workbook  = v_workbook;
    ret_val = t_cnt;
END;
$$;


--
-- Name: add_validation(text, text, text, text, text, text, text, boolean, text, text, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.add_validation(p_type text, p_sqref text, p_style text DEFAULT 'stop'::text, p_formula1 text DEFAULT NULL::text, p_formula2 text DEFAULT NULL::text, p_title text DEFAULT NULL::text, p_prompt text DEFAULT NULL::text, p_show_error boolean DEFAULT false, p_error_title text DEFAULT NULL::text, p_error_txt text DEFAULT NULL::text, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
/* stop, warning, information */
DECLARE
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
    v_validations  pgexcel_generator.tp_validation;
   v_validations_tab  pgexcel_generator.tp_validation[];
	v_sheet pgexcel_generator.tp_sheet;
    t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
BEGIN
    ----PERFORM pgexcel_generator.Init();
    t_ind := COALESCE(array_length(v_workbook.sheets[ t_sheet].validations, 1), 0) + 1;
	v_sheet = v_workbook.sheets[t_sheet];
	v_validations_tab = v_sheet.validations;
	v_validations = v_sheet.validations[t_ind];
     v_validations := v_workbook.sheets[t_sheet].validations[t_ind];
	v_validations.type =  p_type;
    v_validations.errorstyle =  p_style;
    v_validations.sqref =  p_sqref;
    v_validations.formula1 =  p_formula1;
    v_validations.error_title = p_error_title;
    v_validations.error_txt =  p_error_txt;
    v_validations.title =  p_title;
    v_validations.prompt =  p_prompt;
   v_validations.showerrormessage =  p_show_error;
 v_validations_tab[t_ind] := v_validations;
	v_sheet.validations := v_validations_tab;
	 v_workbook.sheets[t_sheet] := v_sheet;
      p_workbook  = v_workbook;
END;
$$;


--
-- Name: alfan_col(integer); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.alfan_col(p_col integer) RETURNS text
    LANGUAGE plpgsql
    AS $$
BEGIN
    RETURN
    CASE
        WHEN p_col > 702 THEN CONCAT_WS('', CHR(TRUNC((64 + TRUNC(((p_col - 27)::NUMERIC / 676::NUMERIC)::NUMERIC))::NUMERIC)::INTEGER), CHR(TRUNC((65 + MOD(TRUNC(((p_col - 1)::NUMERIC / 26::NUMERIC)::NUMERIC) - 1, 26))::NUMERIC)::INTEGER), CHR(TRUNC((65 + MOD(p_col - 1, 26))::NUMERIC)::INTEGER))
        WHEN p_col > 26 THEN CONCAT_WS('', CHR(TRUNC((64 + TRUNC(((p_col - 1)::NUMERIC / 26::NUMERIC)::NUMERIC))::NUMERIC)::INTEGER), CHR(TRUNC((65 + MOD(p_col - 1, 26))::NUMERIC)::INTEGER))
        ELSE CHR(TRUNC((64 + p_col)::NUMERIC)::INTEGER)
    END;
END;
$$;


--
-- Name: blob2file(text, text, text); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.blob2file(p_blob text, p_directory text DEFAULT 'MY_DIR'::text, p_filename text DEFAULT 'my.xlsx'::text)
    LANGUAGE plpgsql
    AS $$
DECLARE
  l_Status boolean;
  vFileHandle record;
  l_file_size numeric;
BEGIN
select * from utl_file_utility.fopen(p_directory, p_filename, 'w' ) into vFileHandle ;
select * from utl_file_utility.put_line( p_directory, p_filename,p_blob ) into l_status ;
select * from utl_file_utility.fclose(p_directory, p_filename  ) into l_file_size ;
END;
$$;


--
-- Name: cell(integer, integer, double precision, integer, integer, integer, integer, pgexcel_generator.tp_alignment, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.cell(p_col integer, p_row integer, p_value double precision, p_numfmtid integer DEFAULT NULL::integer, p_fontid integer DEFAULT NULL::integer, p_fillid integer DEFAULT NULL::integer, p_borderid integer DEFAULT NULL::integer, p_alignment pgexcel_generator.tp_alignment DEFAULT NULL::pgexcel_generator.tp_alignment, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
    t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
	v_rows  pgexcel_generator.tp_cells;
    v_rows_tab  pgexcel_generator.tp_cells[];
	v_sheet pgexcel_generator.tp_sheet;
	v_cell pgexcel_generator.tp_cell;
BEGIN
    				 --raise notice 'start of cell for number';
    --PERFORM pgexcel_generator.Init();
	v_sheet = v_workbook.sheets[t_sheet];
	v_rows =  v_workbook.sheets[t_sheet].rows[p_row];
	v_cell = v_workbook.sheets[t_sheet].rows[p_row].cell[p_col];
	v_cell.value = p_value;
	v_cell.style = null;
	v_rows.cell[p_col] = v_cell;
	v_sheet.rows[p_row] = v_rows;
	v_workbook.sheets[t_sheet] = v_sheet;
--p_workbook  = v_workbook;
    --v_cell.style = pgexcel_generator.get_xfid(t_sheet, p_col, p_row, p_numFmtId, p_fontid, p_fillid, p_borderid, p_alignment,v_workbook);
select * from pgexcel_generator.get_xfid(t_sheet, p_col, p_row, p_numFmtId, p_fontid, p_fillid, p_borderid, p_alignment,v_workbook)
into rec;
v_cell.style = rec.ret_val;
v_workbook = rec.p_workbook;
	v_cell.value = p_value;
   v_sheet = v_workbook.sheets[t_sheet];
	v_rows =  v_workbook.sheets[t_sheet].rows[p_row];
	v_rows.cell[p_col] = v_cell;
	v_sheet.rows[p_row] = v_rows;
	v_workbook.sheets[t_sheet] = v_sheet;
p_workbook  = v_workbook;
END;
$$;


--
-- Name: cell(integer, integer, text, integer, integer, integer, integer, pgexcel_generator.tp_alignment, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.cell(p_col integer, p_row integer, p_value text, p_numfmtid integer DEFAULT NULL::integer, p_fontid integer DEFAULT NULL::integer, p_fillid integer DEFAULT NULL::integer, p_borderid integer DEFAULT NULL::integer, p_alignment pgexcel_generator.tp_alignment DEFAULT NULL::pgexcel_generator.tp_alignment, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;
 v_workbook pgexcel_generator.tp_book  := p_workbook;
    t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
    t_alignment pgexcel_generator.tp_alignment := p_alignment;
	v_rows  pgexcel_generator.tp_cells;
    v_rows_tab  pgexcel_generator.tp_cells[];
	v_sheet pgexcel_generator.tp_sheet;
	v_cell pgexcel_generator.tp_cell;
	add_string_variable int;
	n_xfd_Val text;
BEGIN
--raise notice 'start of cell for text';
    --PERFORM pgexcel_generator.Init();

select * from pgexcel_generator.add_string(p_value,v_workbook) into rec;
add_string_variable := rec.ret_val;
v_workbook := rec.p_workbook;
 v_sheet = v_workbook.sheets[t_sheet];
	v_rows =  v_workbook.sheets[t_sheet].rows[p_row];
	v_cell = v_workbook.sheets[t_sheet].rows[p_row].cell[p_col];
	v_cell.value := add_string_variable;

	v_rows.cell[p_col] = v_cell;
	v_sheet.rows[p_row] = v_rows;
	v_workbook.sheets[t_sheet] = v_sheet;
    IF t_alignment.wrapText IS NULL AND pgexcel_generator.INSTR(p_value, CHR(13)) > 0 THEN
        t_alignment.wrapText := TRUE;
    END IF;
	 v_rows.cell[p_col] = v_cell;
	v_sheet.rows[p_row] = v_rows;
	v_workbook.sheets[t_sheet] = v_sheet;
 p_workbook  = v_workbook;

select * from pgexcel_generator.get_xfid(t_sheet, p_col, p_row, p_numfmtid, p_fontid, p_fillid, p_borderid, t_alignment,v_workbook) into rec;
n_xfd_Val := rec.ret_val;
v_workbook := rec.p_workbook;
      v_cell.style :=  CONCAT_WS('', 't="s" ', n_xfd_Val);
 
     v_rows.cell[p_col] = v_cell;
	v_sheet.rows[p_row] = v_rows;
	v_workbook.sheets[t_sheet] = v_sheet;
	 p_workbook  = v_workbook;
--raise notice 'end of cell for text';
END;
$$;


--
-- Name: cell(integer, integer, timestamp without time zone, integer, integer, integer, integer, pgexcel_generator.tp_alignment, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.cell(p_col integer, p_row integer, p_value timestamp without time zone, p_numfmtid integer DEFAULT NULL::integer, p_fontid integer DEFAULT NULL::integer, p_fillid integer DEFAULT NULL::integer, p_borderid integer DEFAULT NULL::integer, p_alignment pgexcel_generator.tp_alignment DEFAULT NULL::pgexcel_generator.tp_alignment, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;
    t_numFmtId INTEGER := p_numfmtid;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
    t_sheet INTEGER := COALESCE(p_sheet, cOALESCE(array_length(v_workbook.sheets, 1), 0) );
	v_rows  pgexcel_generator.tp_cells;
    v_rows_tab  pgexcel_generator.tp_cells[];
	v_sheet pgexcel_generator.tp_sheet;
	v_cell pgexcel_generator.tp_cell;
BEGIN
--raise notice 'start of cell for date';
    --PERFORM pgexcel_generator.Init();
	v_sheet = v_workbook.sheets[t_sheet];
	v_rows =  v_workbook.sheets[t_sheet].rows[p_row];
	v_cell = v_workbook.sheets[t_sheet].rows[p_row].cell[p_col];
     v_cell.value := (EXTRACT (EPOCH FROM p_value - TO_DATE('01-01-1900'::TEXT, 'DD-MM-YYYY')) / 86400)::NUMERIC;
    v_rows.cell[p_col] = v_cell;
	v_sheet.rows[p_row] = v_rows;
	v_workbook.sheets[t_sheet] = v_sheet;
 --p_workbook  = v_workbook;

    IF t_numFmtId IS NULL AND NOT  (
		(select case when  v_workbook.sheets[t_sheet].col_fmts[ p_col ] is null then false else true end)
     AND v_workbook.sheets[t_sheet].col_fmts[p_col].numFmtId is not NULL)
     AND  not (  (select case when v_workbook.sheets[t_sheet].row_fmts[p_row] is null then false else true end)
      AND v_workbook.sheets[t_sheet].row_fmts[p_row].numFmtId is not NULL)
     THEN
select * from pgexcel_generator.get_numfmt('mm/dd/yyyy'::TEXT,v_workbook) into rec;
     t_numFmtId := rec.ret_val;
		 v_workbook := rec.p_workbook;
    END IF;
 --p_workbook  = v_workbook;
  select * from  pgexcel_generator.get_xfid(t_sheet, p_col, p_row, t_numFmtId, p_fontid, p_fillid, p_borderid, p_alignment,v_workbook) into rec;
v_cell.style = rec.ret_val;
 v_workbook := rec.p_workbook;

v_sheet = v_workbook.sheets[t_sheet];
	v_rows =  v_workbook.sheets[t_sheet].rows[p_row];
  v_rows.cell[p_col] = v_cell;
	v_sheet.rows[p_row] = v_rows;
	v_workbook.sheets[t_sheet] = v_sheet;

  p_workbook  = v_workbook;
--raise notice 'end of cell for text';
END;
$$;


--
-- Name: clear_workbook(pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.clear_workbook(INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
      t_row_ind INTEGER;
	  v_rows  pgexcel_generator.tp_cells;
	v_sheet pgexcel_generator.tp_sheet;
	v_workbook pgexcel_generator.tp_book := p_workbook;
  BEGIN
      ----PERFORM pgexcel_generator.Init();

      for s in 1 .. COALESCE(array_length(v_workbook.sheets, 1), 0)
      loop
	   
         if COALESCE(array_length(v_workbook.sheets[s].rows, 1), 0)   = 0
         then
           t_row_ind  = null;
          else
           t_row_ind := 1;
         end if;

    while t_row_ind is not null
    loop
	v_rows = null;
	 v_sheet.rows[t_row_ind] = v_rows;
	 v_workbook.sheets[s] = v_sheet;
 
	  if  (t_row_ind) != COALESCE(array_length(v_workbook.sheets[s].rows, 1), 0) then
         t_row_ind := t_row_ind+1;
	  else
	     exit;
	  end if;
    end loop;

   v_sheet.rows = null;
    v_sheet.widths = null;
    v_sheet.autofilters = null;
    v_sheet.hyperlinks = null;
    v_sheet.col_fmts = null;
    v_sheet.row_fmts = null;
    v_sheet.comments = null;
    v_sheet.mergecells  = null;
    v_sheet.validations = null;
	v_workbook.sheets[s] = v_sheet;
  end loop;
  v_workbook.strings = null;
  v_workbook.str_ind = null;
  v_workbook.fonts = null;
  v_workbook.fills = null;
  v_workbook.borders = null;
  v_workbook.numFmts = null;
  v_workbook.cellXfs = null;
  v_workbook.defined_names = null;
  v_workbook := null;

p_workbook  = v_workbook;
  END;
$$;


--
-- Name: col_alfan(text); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.col_alfan(p_col text) RETURNS integer
    LANGUAGE plpgsql
    AS $$
  BEGIN
      RETURN ASCII(pgexcel_generator.substr(p_col, - 1)) - 64 + COALESCE((ASCII(pgexcel_generator.substr(p_col, - 2, 1)) - 64) * 26, 0) + COALESCE((ASCII(pgexcel_generator.substr(p_col, - 3, 1)) - 64) * 676, 0);
  END;
  $$;


--
-- Name: comment(integer, integer, text, text, integer, integer, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.comment(p_col integer, p_row integer, p_text text, p_author text DEFAULT NULL::text, p_width integer DEFAULT 150, p_height integer DEFAULT 100, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
    t_ind INTEGER;
   v_workbook pgexcel_generator.tp_book :=p_workbook;
    t_sheet INTEGER:= COALESCE(p_sheet,  COALESCE(array_length(v_workbook.sheets, 1), 0));
	v_sheet pgexcel_generator.tp_sheet;
	v_comments pgexcel_generator.tp_comment;
BEGIN
    ----PERFORM pgexcel_generator.Init();

    t_ind :=  COALESCE(array_length(v_workbook.sheets[ t_sheet].comments, 1), 0) + 1;
	v_sheet = v_workbook.sheets[t_sheet];
	v_comments =  v_workbook.sheets[t_sheet].comments[t_ind];
    v_comments.row = p_row;
    v_comments.column = p_col;
    v_comments.text =  (p_text);
    v_comments.author =  (p_author);
    v_comments.width = p_width;
    v_comments.height =  p_height;
	 v_sheet.comments[t_ind] = v_comments;
	v_workbook.sheets[t_sheet] = v_sheet;
    p_workbook  = v_workbook;
END;
$$;


--
-- Name: defined_name(integer, integer, integer, integer, text, integer, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.defined_name(p_tl_col integer, p_tl_row integer, p_br_col integer, p_br_row integer, p_name text, p_sheet integer DEFAULT NULL::integer, p_localsheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $_$
/* top left */
/* bottom right */
DECLARE
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
    t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
	v_defined_names pgexcel_generator.tp_defined_name;
BEGIN
    ----PERFORM pgexcel_generator.Init();
    t_ind := COALESCE(array_length(v_workbook.defined_names, 1), 0) + 1;
	v_defined_names = v_workbook.defined_names[t_ind];
    v_defined_names.name = p_name;
    v_defined_names.ref =  CONCAT_WS('', 'Sheet', t_sheet, '!$', pgexcel_generator.alfan_col(p_tl_col), '$', p_tl_row, ':$', pgexcel_generator.alfan_col(p_br_col), '$', p_br_row);
    v_defined_names.sheet =  p_localsheet;
    v_workbook.defined_names[t_ind] = v_defined_names;
    p_workbook  = v_workbook;
END;
$_$;


--
-- Name: finish(text, text, pgexcel_generator.tp_book, pgexcel_generator.tp_authors[]); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.finish(p_directory text, p_filename text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, INOUT p_authors pgexcel_generator.tp_authors[] DEFAULT NULL::pgexcel_generator.tp_authors[], OUT ret_val text) RETURNS record
    LANGUAGE plpgsql
    AS $_$
DECLARE
rec record;
    t_folder CHARACTER VARYING(4000) := p_filename;
    t_excel text;
    t_xxx TEXT;
    t_tmp CHARACTER VARYING;
    t_str CHARACTER VARYING;
    t_c DOUBLE PRECISION;
    t_h DOUBLE PRECISION;
    t_w DOUBLE PRECISION;
    t_cw DOUBLE PRECISION;
    t_cell CHARACTER VARYING(1000);
    t_row_ind INTEGER;
    t_col_min INTEGER;
    t_col_max INTEGER;
    t_col_ind INTEGER;
    t_len INTEGER;
    ts TIMESTAMP(6) WITHOUT TIME ZONE := current_timestamp;
    v_workbook pgexcel_generator.tp_book := p_workbook;
     v_authors pgexcel_generator.tp_authors[] :=  p_authors;
      v_authors_rec pgexcel_generator.tp_authors;
BEGIN
    --PERFORM pgexcel_generator.Init();
    SET client_encoding='UTF8';
	--PERFORM PKG_REBATE_ACCRUAL.Init();
    t_excel := NULL::TEXT;
	t_folder := replace (t_folder, '.xlsx','');
	 -- raise notice 'start finish';
	 -- execute 'create temp table if not exists finish_xls(col text)';
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'::TEXT;
    FOR s IN 1..COALESCE(array_length(v_workbook.sheets, 1), 0) LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<Override PartName="/xl/worksheets/sheet', s, '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'))::TEXT;
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'))::TEXT;
    FOR s IN 1..COALESCE(array_length(v_workbook.sheets, 1), 0) LOOP
        IF COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) > 0 THEN
            t_xxx := (CONCAT_WS('', t_xxx, '<Override PartName="/xl/comments', s, '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>'))::TEXT;
        END IF;
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, '</Types>'))::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, '[Content_Types].xml'::TEXT, t_xxx);
	--insert into finish_xls values(t_xxx);
		  call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','[Content_Types].xml'));
    t_xxx := (CONCAT_WS('', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>', current_user, '</dc:creator><cp:lastModifiedBy>', current_user, '</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">', TO_CHAR(clock_timestamp(), 'yyyy-mm-dd"T"hh24:mi:ss'), '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">', TO_CHAR(clock_timestamp(), 'yyyy-mm-dd"T"hh24:mi:ss'), '</dcterms:modified></cp:coreProperties>'))::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, 'docProps/core.xml'::TEXT, t_xxx);
	--insert into finish_xls values(t_xxx);
	 call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','docProps/core.xml'));
    t_xxx := (CONCAT_WS('', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>', COALESCE(array_length(v_workbook.sheets, 1), 0), '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="', COALESCE(array_length(v_workbook.sheets, 1), 0), '" baseType="lpstr">'))::TEXT;
    FOR s IN 1..COALESCE(array_length(v_workbook.sheets, 1), 0) LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<vt:lpstr>',  v_workbook.sheets[s].name, '</vt:lpstr>'))::TEXT;
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, '</vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0300</AppVersion></Properties>'))::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, 'docProps/app.xml'::TEXT, t_xxx);
	--insert into finish_xls values(t_xxx);
	 call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','docProps/app.xml'));
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, '_rels/.rels'::TEXT, t_xxx);
        --insert into finish_xls values(t_xxx);
		 call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','_rels/.rels'));
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'::TEXT;
   IF COALESCE(array_length(v_workbook.numFmts, 1), 0) > 0 THEN
        t_xxx := (CONCAT_WS('', t_xxx, '<numFmts count="', COALESCE(array_length(v_workbook.numFmts, 1), 0), '">'))::TEXT;
        FOR n IN 1..COALESCE(array_length(v_workbook.numFmts, 1), 0) LOOP
            t_xxx := (CONCAT_WS('', t_xxx, '<numFmt numFmtId="', v_workbook.numFmts[n].numFmtId, '" formatCode="', v_workbook.numFmts[n].formatCode, '"/>'))::TEXT;
        END LOOP;
        t_xxx := (CONCAT_WS('', t_xxx, '</numFmts>'))::TEXT;
    END IF;

    t_xxx := (CONCAT_WS('', t_xxx, '<fonts count="', COALESCE(array_length(v_workbook.fonts, 1), 0), '" x14ac:knownFonts="1">'))::TEXT;
    FOR f IN 0..COALESCE(array_length(v_workbook.fonts, 1), 0) - 1 LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<font>',
        CASE
            WHEN v_workbook.fonts[f].bold THEN '<b/>'
        END,
        CASE
            WHEN v_workbook.fonts[f].italic THEN '<i/>'
        END,
        CASE
            WHEN v_workbook.fonts[f].underline THEN '<u/>'
        END, '<sz val="', v_workbook.fonts[f].fontsize, '"/><color ',
        CASE
            WHEN v_workbook.fonts[f].rgb IS NOT NULL THEN CONCAT_WS('', 'rgb="', v_workbook.fonts[f].rgb)
            ELSE CONCAT_WS('', 'theme="', v_workbook.fonts[f].theme)
        END, '"/><name val="', v_workbook.fonts[f].name, '"/><family val="', v_workbook.fonts[f].family, '"/><scheme val="none"/></font>'))::TEXT;
    END LOOP;
	
    t_xxx := (CONCAT_WS('', t_xxx, '</fonts><fills count="', COALESCE(array_length(v_workbook.fills, 1), 0), '">'))::TEXT;
    FOR f IN 0..COALESCE(array_length(v_workbook.fills, 1), 0) - 1 LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<fill><patternFill patternType="',  v_workbook.fills[f].patternType, '">',
        CASE
            WHEN  v_workbook.fills[f].fgRGB IS NOT NULL THEN CONCAT_WS('', '<fgColor rgb="',  v_workbook.fills[f].fgRGB, '"/>')
        END, '</patternFill></fill>'))::TEXT;
    END LOOP;

    t_xxx := (CONCAT_WS('', t_xxx, '</fills><borders count="', COALESCE(array_length(v_workbook.borders, 1), 0), '">'))::TEXT;
    FOR b IN 0..COALESCE(array_length(v_workbook.borders, 1), 0) - 1 LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<border>',
        CASE
            WHEN v_workbook.borders[b].left IS NULL THEN '<left/>'
            ELSE CONCAT_WS('', '<left style="', v_workbook.borders[b].left, '"/>')
        END,
        CASE
            WHEN v_workbook.borders[b].right IS NULL THEN '<right/>'
            ELSE CONCAT_WS('', '<right style="', v_workbook.borders[b].right, '"/>')
        END,
        CASE
            WHEN v_workbook.borders[b].top IS NULL THEN '<top/>'
            ELSE CONCAT_WS('', '<top style="', v_workbook.borders[b].top, '"/>')
        END,
        CASE
            WHEN v_workbook.borders[b].bottom IS NULL THEN '<bottom/>'
            ELSE CONCAT_WS('', '<bottom style="', v_workbook.borders[b].bottom, '"/>')
        END, '</border>'))::TEXT;
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="', (COALESCE(array_length(v_workbook.cellXfs, 1), 0)  + 1), '"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'))::TEXT;

    FOR x IN 1..COALESCE(array_length(v_workbook.cellXfs, 1), 0)  LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<xf numFmtId="', v_workbook.cellXfs[x].numFmtId, '" fontId="', v_workbook.cellXfs[x].fontId, '" fillId="', v_workbook.cellXfs[x].fillId, '" borderId="', v_workbook.cellXfs[x].borderId, '">'))::TEXT;
        IF (v_workbook.cellXfs[x].alignment.horizontal IS NOT NULL OR v_workbook.cellXfs[x].alignment.vertical IS NOT NULL OR v_workbook.cellXfs[x].alignment.wrapText) THEN
            t_xxx := (CONCAT_WS('', t_xxx, '<alignment',
            CASE
                WHEN v_workbook.cellXfs[x].alignment.horizontal IS NOT NULL THEN CONCAT_WS('', ' horizontal="', v_workbook.cellXfs[x].alignment.horizontal, '"')
            END,
            CASE
                WHEN v_workbook.cellXfs[x].alignment.vertical IS NOT NULL THEN CONCAT_WS('', ' vertical="', v_workbook.cellXfs[x].alignment.vertical, '"')
            END,
            CASE
                WHEN v_workbook.cellXfs[x].alignment.wrapText THEN ' wrapText="true"'
            END, '/>'))::TEXT;
        END IF;
        t_xxx := (CONCAT_WS('', t_xxx, '</xf>'))::TEXT;
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>'))::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, 'xl/styles.xml'::TEXT, t_xxx);
	--insert into finish_xls values(t_xxx);
	 call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','xl/styles.xml'));
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/><workbookPr date1904="false" defaultThemeVersion="124226"/><bookViews><workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/></bookViews><sheets>'::TEXT;
 
    FOR s IN 1..COALESCE(array_length(v_workbook.sheets, 1), 0) LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<sheet name="', v_workbook.sheets[s].name, '" sheetId="', s, '" r:id="rId', (9 + s), '"/>'))::TEXT;
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, '</sheets>'))::TEXT;
    IF COALESCE(array_length(v_workbook.defined_names, 1), 0) > 0 THEN
        t_xxx := (CONCAT_WS('', t_xxx, '<definedNames>'))::TEXT;
        FOR s IN 1..COALESCE(array_length(v_workbook.defined_names, 1), 0) LOOP
            t_xxx := (CONCAT_WS('', t_xxx, '
<definedName name="',  v_workbook.defined_names[s].name, '"',
            CASE
                WHEN  v_workbook.defined_names[s].sheet IS NOT NULL THEN CONCAT_WS('', ' localSheetId="',  v_workbook.defined_names[s].sheet, '"')
            END, '>',  v_workbook.defined_names[s].ref, '</definedName>'))::TEXT;
        END LOOP;
        t_xxx := (CONCAT_WS('', t_xxx, '</definedNames>'))::TEXT;
    END IF;
    t_xxx := (CONCAT_WS('', t_xxx, '<calcPr calcId="144525"/></workbook>'))::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, 'xl/workbook.xml'::TEXT, t_xxx);
	--insert into finish_xls values(t_xxx);
	call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','xl/workbook.xml'));
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, 'xl/theme/theme1.xml'::TEXT, t_xxx);
    --insert into finish_xls values(t_xxx);
	call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','xl/theme/theme1.xml'));
 
    FOR s IN 1..COALESCE(array_length(v_workbook.sheets, 1), 0) LOOP
        t_col_min := 16384;
        t_col_max := 1;

        t_row_ind :=  case  when array_length((v_workbook.sheets[s].rows)::pgexcel_generator.tp_cells[], 1) is null then null  else 1 end ;

        WHILE t_row_ind <= COALESCE(array_upper((v_workbook.sheets[s].rows[t_row_ind].cell)::pgexcel_generator.tp_cell[],1),1)  LOOP
         --  WHILE t_row_ind <= array_upper((v_workbook.sheets[s].rows)::pgexcel_generator.tp_cells[],1)  LOOP
            t_col_min := least(t_col_min,  COALESCE(case  when array_length((v_workbook.sheets[s].rows[t_row_ind].cell)::pgexcel_generator.tp_cell[], 1) is null then null  else 1 end,1));

            t_col_max := greatest(t_col_max,  COALESCE(case  when array_length((v_workbook.sheets[s].rows[t_row_ind].cell)::pgexcel_generator.tp_cell[],1) is null then null  else array_upper((v_workbook.sheets[s].rows[t_row_ind].cell)::pgexcel_generator.tp_cell[],1) end,0));

            t_row_ind := t_row_ind+1;
        END LOOP;

        t_xxx := (CONCAT_WS('', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="', pgexcel_generator.alfan_col(t_col_min),  case  when array_length((v_workbook.sheets[s].rows)::pgexcel_generator.tp_cells[], 1) is null then null  else 1 end, ':', pgexcel_generator.alfan_col(t_col_max), array_length((v_workbook.sheets[s].rows)::pgexcel_generator.tp_cells[], 1), '"/><sheetViews><sheetView',
        CASE
            WHEN s = 1 THEN ' tabSelected="1"'
        END, ' workbookViewId="0">'))::TEXT;

        IF v_workbook.sheets[s].freeze_rows > 0 AND v_workbook.sheets[s].freeze_cols > 0 THEN
            t_xxx := (CONCAT_WS('', t_xxx, (CONCAT_WS('', '<pane xSplit="', v_workbook.sheets[s].freeze_cols, '" ', 'ySplit="', v_workbook.sheets[s].freeze_rows, '" ', 'topLeftCell="', pgexcel_generator.alfan_col(v_workbook.sheets[s].freeze_cols + 1), (v_workbook.sheets[s].freeze_rows + 1), '" ', 'activePane="bottomLeft" state="frozen"/>'))))::TEXT;
        ELSE
            IF v_workbook.sheets[s].freeze_rows > 0 THEN
                t_xxx := (CONCAT_WS('', t_xxx, '<pane ySplit="', v_workbook.sheets[s].freeze_rows, '" topLeftCell="A', (v_workbook.sheets[s].freeze_rows + 1), '" activePane="bottomLeft" state="frozen"/>'))::TEXT;
            END IF;
            IF v_workbook.sheets[s].freeze_cols > 0 THEN
                t_xxx := (CONCAT_WS('', t_xxx, '<pane xSplit="', v_workbook.sheets[s].freeze_cols, '" topLeftCell="', pgexcel_generator.alfan_col(v_workbook.sheets[s].freeze_cols + 1), '1" activePane="bottomLeft" state="frozen"/>'))::TEXT;
            END IF;
        END IF;

        t_xxx := (CONCAT_WS('', t_xxx, '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'))::TEXT;
        IF COALESCE(array_length(v_workbook.sheets[s].widths, 1), 0) > 0 THEN
            t_xxx := (CONCAT_WS('', t_xxx, '<cols>'))::TEXT;
            t_col_ind := 1;
            WHILE t_col_ind <= array_length(v_workbook.sheets[s].widths, 1) LOOP
			
			--IF v_workbook.sheets[s].widths[t_col_ind] is not null THEN
                t_xxx := concat_ws('', t_xxx , '<col min="' , t_col_ind ,'" max="' , t_col_ind , '" width="' ,  COALESCE(v_workbook.sheets[s].widths[t_col_ind],10) , '" customWidth="1"/>');
			--END IF;
			
                   t_col_ind :=  t_col_ind+1;
			
            END LOOP;
            t_xxx := (CONCAT_WS('', t_xxx, '</cols>'))::TEXT;
        END IF;
        t_xxx := (CONCAT_WS('', t_xxx, '<sheetData>'))::TEXT;
        t_row_ind :=case when array_length((v_workbook.sheets[s].rows)::pgexcel_generator.tp_cells[], 1) is null then null else 1 end;

        t_tmp := NULL;
  
	  
	  -- raise notice 'spans t_col_min- %, t_col_max -%',t_col_min,t_col_max;
	  WHILE t_row_ind <= array_upper((v_workbook.sheets[s].rows)::pgexcel_generator.tp_cells[], 1) LOOP
            t_tmp := CONCAT_WS('', t_tmp, '<row r="', t_row_ind, '" spans="', t_col_min, ':', t_col_max, '">');
            t_len := LENGTH(t_tmp);
            t_col_ind := case when array_length((v_workbook.sheets[s].rows[t_row_ind].cell)::pgexcel_generator.tp_cell[],1) is null
			then null else 1 end;

		
			  --WHILE t_col_ind <= array_length(v_workbook.sheets[s].rows[t_row_ind].cell::pgexcel_generator.tp_cell[],1) LOOP
			  WHILE t_col_ind <= array_upper(v_workbook.sheets[s].rows[t_row_ind].cell::pgexcel_generator.tp_cell[],1) LOOP

            t_cell := CONCAT_WS('', '<c r="', pgexcel_generator.alfan_col(t_col_ind), t_row_ind, '"', ' ', v_workbook.sheets[s].rows[t_row_ind].cell[t_col_ind ].style, '><v>',  v_workbook.sheets[ s ].rows[t_row_ind].cell[t_col_ind].value, '</v></c>');

 /*               IF t_len > 32000 THEN

                    t_tmp := NULL;
                     t_len := 0;
                END IF; */
             t_tmp := CONCAT_WS('', t_tmp, t_cell);
             t_len := t_len + LENGTH(t_cell);
             t_col_ind :=  t_col_ind+1 ;
            END LOOP;
           t_tmp := CONCAT_WS('', t_tmp, '</row>');
           t_row_ind := t_row_ind+1;
        END LOOP;
        t_tmp := CONCAT_WS('', t_tmp, '</sheetData>');
        t_len := LENGTH(t_tmp);
        /*
        [5340 - Severity CRITICAL - PostgreSQL doesn't support the SYS.DBMS_LOB.WRITEAPPEND(CLOB,INTEGER,VARCHAR2) function. Use suitable function or create user defined function.]
        dbms_lob.writeappend( t_xxx, t_len, t_tmp )
        */
         t_xxx = concat_ws('',t_xxx,'',t_tmp,'');
        FOR a IN 1..COALESCE(array_length(v_workbook.sheets[s].autofilters, 1), 0) LOOP
            t_xxx := (CONCAT_WS('', t_xxx, '<autoFilter ref="', pgexcel_generator.alfan_col(COALESCE(v_workbook.sheets[s].autofilters[a].column_start, t_col_min)), COALESCE(v_workbook.sheets[s].autofilters[a].row_start, 1 ), ':', pgexcel_generator.alfan_col(COALESCE(v_workbook.sheets[s].autofilters[a].column_end, v_workbook.sheets[s].autofilters[a].column_start, t_col_max)), COALESCE(v_workbook.sheets[s].autofilters[a].row_end, array_length((v_workbook.sheets[s].rows)::pgexcel_generator.tp_cells[], 1)), '"/>')::TEXT);

        END LOOP;
        IF COALESCE(array_length(v_workbook.sheets[s].mergecells, 1), 0) > 0 THEN
            t_xxx := (CONCAT_WS('', t_xxx, '<mergeCells count="', COALESCE(array_length(v_workbook.sheets[s].mergecells, 1), 0), '">'));
            FOR m IN 1..COALESCE(array_length(v_workbook.sheets[s].mergecells, 1), 0) LOOP
                BEGIN
                    t_xxx :=concat_ws('', t_xxx ,'<mergeCell ref="' , v_workbook.sheets[ s ].mergecells[ m ] , '"/>');
                END;
            END LOOP;
            t_xxx := (CONCAT_WS('', t_xxx, '</mergeCells>'))::TEXT;
        END IF;
        /* */

        IF COALESCE(array_length(v_workbook.sheets[s].validations, 1), 0)> 0 THEN
            t_xxx := (CONCAT_WS('', t_xxx, '<dataValidations count="', COALESCE(array_length(v_workbook.sheets[s].validations, 1), 0), '">'));
            FOR m IN 1..COALESCE(array_length(v_workbook.sheets[s].validations, 1), 0)LOOP
                t_xxx := (CONCAT_WS('', t_xxx, '<dataValidation', ' type="', v_workbook.sheets[s].validations.type, '"', ' errorStyle="', v_workbook.sheets[s].validations.errorstyle, '"', ' allowBlank="',
                CASE
                    WHEN COALESCE(v_workbook.sheets[s].validations.allowBlank, TRUE) THEN '1'
                    ELSE '0'
                END, '"', ' sqref="', v_workbook.sheets[s].validations.sqref, '"'))::TEXT;
                IF v_workbook.sheets[s].validations.prompt IS NOT NULL THEN
                    t_xxx := (CONCAT_WS('', t_xxx, ' showInputMessage="1" prompt="', v_workbook.sheets[s].validations.prompt, '"'))::TEXT;
                    IF v_workbook.sheets[s].validations.title IS NOT NULL THEN
                        t_xxx := (CONCAT_WS('', t_xxx, ' promptTitle="', v_workbook.sheets[s].validations.title, '"'))::TEXT;
                    END IF;
                END IF;
                IF v_workbook.sheets[s].validations.showerrormessage THEN
                    t_xxx := (CONCAT_WS('', t_xxx, ' showErrorMessage="1"'))::TEXT;
                    IF v_workbook.sheets[s].validations.error_title IS NOT NULL THEN
                        t_xxx := (CONCAT_WS('', t_xxx, ' errorTitle="', v_workbook.sheets[s].validations.error_title, '"'))::TEXT;
                    END IF;
                    IF v_workbook.sheets[s].validations.error_txt IS NOT NULL THEN
                        t_xxx := (CONCAT_WS('', t_xxx, ' error="', v_workbook.sheets[s].validations.error_txt, '"'))::TEXT;
                    END IF;
                END IF;
                t_xxx := (CONCAT_WS('', t_xxx, '>'))::TEXT;
                IF v_workbook.sheets[s].validations.formula1 IS NOT NULL THEN
                    t_xxx := (CONCAT_WS('', t_xxx, '<formula1>', v_workbook.sheets[s].validations.formula1, '</formula1>'))::TEXT;
                END IF;
                IF v_workbook.sheets[s].validations.formula2 IS NOT NULL THEN
                    t_xxx := (CONCAT_WS('', t_xxx, '<formula2>', v_workbook.sheets[s].validations.formula2, '</formula2>'))::TEXT;
                END IF;
                t_xxx := (CONCAT_WS('', t_xxx, '</dataValidation>'))::TEXT;
            END LOOP;
            t_xxx := (CONCAT_WS('', t_xxx, '</dataValidations>'))::TEXT;
        END IF;
        /* */

        IF COALESCE(array_length(v_workbook.sheets[s].hyperlinks, 1), 0) > 0 THEN
            t_xxx := (CONCAT_WS('', t_xxx, '<hyperlinks>'))::TEXT;
            FOR h IN 1..COALESCE(array_length(v_workbook.sheets[s].hyperlinks, 1), 0) LOOP
                t_xxx := (CONCAT_WS('', t_xxx, '<hyperlink ref="', v_workbook.sheets[s].hyperlinks.cell, '" r:id="rId', h, '"/>'))::TEXT;
            END LOOP;
            t_xxx := (CONCAT_WS('', t_xxx, '</hyperlinks>'))::TEXT;
        END IF;
        t_xxx := (CONCAT_WS('', t_xxx, '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'))::TEXT;
        IF COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) > 0 THEN
            t_xxx := (CONCAT_WS('', t_xxx, '<legacyDrawing r:id="rId', COALESCE(array_length(v_workbook.sheets[s].hyperlinks, 1), 0) + 1, '"/>'))::TEXT;
        END IF;
        /* */
        t_xxx := (CONCAT_WS('', t_xxx, '</worksheet>'))::TEXT;
        --CALL pgexcel_generator.add1xml(t_excel, CONCAT_WS('', 'xl/worksheets/sheet', s, '.xml'), t_xxx);
          --insert into finish_xls values(t_xxx);
		  call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','xl/worksheets/sheet',s,'.xml'));
        IF COALESCE(array_length(v_workbook.sheets[s].hyperlinks, 1), 0) > 0 OR COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) > 0 THEN
            t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'::TEXT;
            IF COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) > 0 THEN
                t_xxx := (CONCAT_WS('', t_xxx, '<Relationship Id="rId', (COALESCE(array_length(v_workbook.sheets[s].hyperlinks, 1), 0) + 2), '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments', s, '.xml"/>'))::TEXT;
                t_xxx := (CONCAT_WS('', t_xxx, '<Relationship Id="rId', (COALESCE(array_length(v_workbook.sheets[s].hyperlinks, 1), 0) + 1), '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing', s, '.vml"/>'))::TEXT;
            END IF;
            FOR h IN 1..COALESCE(array_length(v_workbook.sheets[s].hyperlinks, 1), 0) LOOP
                t_xxx := (CONCAT_WS('', t_xxx, '<Relationship Id="rId', h, '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="', v_workbook.sheets[s].hyperlinks.url, '" TargetMode="External"/>'))::TEXT;
            END LOOP;
            t_xxx := (CONCAT_WS('', t_xxx, '</Relationships>'))::TEXT;
            --CALL pgexcel_generator.add1xml(t_excel, CONCAT_WS('', 'xl/worksheets/_rels/sheet', s, '.xml.rels'), t_xxx);
              ---insert into finish_xls values(t_xxx);
			  call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','xl/worksheets/_rels/sheet',s,'.xml.rels'));
        END IF;
        /* */

        IF COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) > 0 THEN
            DECLARE
--rec record;
                cnt INTEGER;
                author_ind CHARACTER VARYING;
            /* t_col_ind := workbook.sheets( s ).widths.next( t_col_ind ); */
            BEGIN
			   v_authors = null;
                FOR c IN 1..COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) LOOP
                   --  v_authors[ v_workbook.sheets[ s ].comments[ c ].author ] := 0;
				   v_authors_rec.col_ind = v_workbook.sheets[ s ].comments[ c ].author;
				   v_authors_rec.column_value = 0;
				   v_authors[c] = v_authors_rec;
                END LOOP;
                t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>'::TEXT;
                cnt := 0;
                author_ind := case when  array_length(v_authors,1) is null then null else 1 end;
                WHILE author_ind <= array_length(v_authors,1) OR  v_authors[author_ind+1] IS NOT NULL LOOP
                      v_authors[author_ind] := cnt;
                       t_xxx := (CONCAT_WS('', t_xxx, '<author>', author_ind, '</author>'))::TEXT;
                         cnt := cnt + 1;
                    author_ind := v_authors[author_ind+1];
                END LOOP;
			p_authors := v_authors;
            END;
            t_xxx := (CONCAT_WS('', t_xxx, '</authors><commentList>'))::TEXT;
            FOR c IN 1..COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) LOOP
                t_xxx := CONCAT_WS('', t_xxx, '<comment ref="', pgexcel_generator.alfan_col(v_workbook.sheets[s].comments.column), CONCAT_WS('', v_workbook.sheets[s].comments.row, '" authorId="', v_authors[workbook.sheets[s].comments[ c ].author] ), '"><text>');
                IF v_workbook.sheets[s].comments.author IS NOT NULL THEN
                    t_xxx := (CONCAT_WS('', t_xxx, '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">', v_workbook.sheets[s].comments.author, ':</t></r>'))::TEXT;
                END IF;
                t_xxx := (CONCAT_WS('', t_xxx, '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">',
                CASE
                    WHEN v_workbook.sheets[s].comments.author IS NOT NULL THEN '
'
                END, v_workbook.sheets[s].comments.text, '</t></r></text></comment>'))::TEXT;
            END LOOP;
            t_xxx := (CONCAT_WS('', t_xxx, '</commentList></comments>'))::TEXT;
            --CALL pgexcel_generator.add1xml(t_excel, CONCAT_WS('','t_folder', 'xl/comments', s, '.xml'), t_xxx);
              --insert into finish_xls values(t_xxx);
			   call pgexcel_generator.blob2file(t_xxx, p_directory, CONCAT_WS('',t_folder,'/', 'xl/comments', s, '.xml'));
            t_xxx := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
			<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'::TEXT;
            FOR c IN 1..COALESCE(array_length(v_workbook.sheets[s].comments, 1), 0) LOOP
                t_xxx := (CONCAT_WS('', t_xxx, '<v:shape id="_x0000_s', TO_CHAR(c,'999'), '" type="#_x0000_t202" style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:', c, ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'))::TEXT;
                t_w := v_workbook.sheets[s].comments.width;
                t_c := 1;
                LOOP
                    IF v_workbook.sheets[s].widths[workbook.sheets[s].comments[c].column + t_c ] is not null THEN
                          t_cw := 256 * v_workbook.sheets[s].widths[workbook.sheets[s].comments[c].column + t_c ];
                        t_cw := TRUNC(((t_cw + 18)::NUMERIC / 256::NUMERIC * 7)::NUMERIC);
                    /* assume default 11 point Calibri */
                    ELSE
                        t_cw := 64;
                    END IF;
                    EXIT WHEN t_w < t_cw;
                    t_c := t_c + 1;
                    t_w := t_w - t_cw;
                END LOOP;
                t_h := v_workbook.sheets[s].comments.height;
                t_xxx := (CONCAT_WS('', t_xxx, CONCAT_WS('', '<x:Anchor>', v_workbook.sheets[s].comments.column, ',15,', v_workbook.sheets[s].comments.row, ',30,', (v_workbook.sheets[s].comments.column + t_c - 1), ',', ROUND(t_w), ',', (v_workbook.sheets[s].comments.row + 1 + TRUNC((t_h / 20::NUMERIC)::NUMERIC)), ',', MOD(t_h, 20), '</x:Anchor>')));
                t_xxx := (CONCAT_WS('', t_xxx, CONCAT_WS('', '<x:AutoFill>False</x:AutoFill><x:Row>', (v_workbook.sheets[s].comments.row - 1), '</x:Row><x:Column>', (v_workbook.sheets[s].comments.column - 1), '</x:Column></x:ClientData></v:shape>')));
            END LOOP;
            t_xxx := (CONCAT_WS('', t_xxx, '</xml>'))::TEXT;
            --CALL pgexcel_generator.add1xml(t_excel, CONCAT_WS('', 'xl/drawings/vmlDrawing', s, '.vml'), t_xxx);
  ---insert into finish_xls values(t_xxx);
   call pgexcel_generator.blob2file(t_xxx, p_directory, CONCAT_WS('',t_folder,'/', 'xl/drawings/vmlDrawing', s, '.vml'));
        END IF;
    /* */
    END LOOP;

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'::TEXT;
    FOR s IN 1..COALESCE(array_length(v_workbook.sheets, 1), 0) LOOP
        t_xxx := (CONCAT_WS('', t_xxx, '<Relationship Id="rId', (9 + s), '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet', s, '.xml"/>'))::TEXT;
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, '</Relationships>'))::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, 'xl/_rels/workbook.xml.rels'::TEXT, t_xxx);
	--insert into finish_xls values(t_xxx);
	   call pgexcel_generator.blob2file(t_xxx, p_directory, CONCAT_WS('',t_folder,'/', 'xl/_rels/workbook.xml.rels'));
    t_xxx := (CONCAT_WS('', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="', v_workbook.str_cnt, '" uniqueCount="', COALESCE(array_length(v_workbook.strings, 1), 0), '">'))::TEXT;
    t_tmp := NULL;
    FOR i IN 0..COALESCE(array_length(v_workbook.str_ind, 1), 0) - 1 LOOP
        t_str := concat_ws('','<si><t>' , substr(v_workbook.str_ind[i], 1, 32000 ) , '</t></si>');

        IF LENGTH(t_tmp) + LENGTH(t_str) > 32000 THEN
            t_xxx := (CONCAT_WS('', t_xxx, t_tmp))::TEXT;
            t_tmp := NULL;
        END IF;
        t_tmp := CONCAT_WS('', t_tmp, t_str);
    END LOOP;
    t_xxx := (CONCAT_WS('', t_xxx, t_tmp, '</sst>'))::TEXT;
    --CALL pgexcel_generator.add1xml(t_excel, 'xl/sharedStrings.xml'::TEXT, t_xxx);
	--insert into finish_xls values(t_xxx);
	call pgexcel_generator.blob2file(t_xxx, p_directory, concat_ws('',t_folder,'/','xl/sharedStrings.xml'));
   -- CALL pgexcel_generator.finish_zip(t_excel);

    CALL pgexcel_generator.clear_workbook(v_workbook);

      p_workbook  = v_workbook;
	 -- select string_agg(col,'') from  finish_xls into t_excel;
t_excel := 'Finish successful';
--raise notice 'end of finish function: %', clock_timestamp();
    ret_val = t_excel;
END;
$_$;


--
-- Name: freeze_cols(integer, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.freeze_cols(p_nr_cols integer DEFAULT 1, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
v_workbook pgexcel_generator.tp_book := p_workbook;
 t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
v_sheet pgexcel_generator.tp_sheet;
  BEGIN
    ----PERFORM pgexcel_generator.Init();
	v_sheet = v_workbook.sheets[ t_sheet ];
    v_sheet.freeze_cols := p_nr_cols;
     v_sheet.freeze_rows := null;
	 v_workbook.sheets[ t_sheet ] = v_sheet;
        p_workbook  = v_workbook;
END;
$$;


--
-- Name: freeze_pane(integer, integer, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.freeze_pane(p_col integer, p_row integer, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
v_workbook pgexcel_generator.tp_book := p_workbook;
 t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
v_sheet pgexcel_generator.tp_sheet;
BEGIN
    ----PERFORM pgexcel_generator.Init();
    v_sheet = v_workbook.sheets[ t_sheet ];
    v_sheet.freeze_cols := p_col;
    v_sheet.freeze_rows := p_row;
    v_workbook.sheets[ t_sheet ] = v_sheet;
      p_workbook  = v_workbook;
END;
$$;


--
-- Name: freeze_rows(integer, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.freeze_rows(p_nr_rows integer DEFAULT 1, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
declare
v_workbook pgexcel_generator.tp_book  := p_workbook;
 t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
 v_sheet pgexcel_generator.tp_sheet;
  BEGIN
    ----PERFORM pgexcel_generator.Init();
    v_sheet = v_workbook.sheets[ t_sheet ];
    v_sheet.freeze_cols := null;
    v_sheet.freeze_rows := p_nr_rows;
    v_workbook.sheets[ t_sheet ] = v_sheet;
      p_workbook  = v_workbook;
    END;
$$;


--
-- Name: get_alignment(text, text, boolean); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.get_alignment(p_vertical text DEFAULT NULL::text, p_horizontal text DEFAULT NULL::text, p_wraptext boolean DEFAULT NULL::boolean) RETURNS pgexcel_generator.tp_alignment
    LANGUAGE plpgsql
    AS $$
DECLARE
    t_rv pgexcel_generator.tp_alignment ;
BEGIN
t_rv.vertical := p_vertical;
t_rv.horizontal := p_horizontal;
 t_rv.wrapText := p_wrapText;
    RETURN t_rv;
END;
$$;


--
-- Name: get_border(text, text, text, text, pgexcel_generator.tp_book); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.get_border(p_top text DEFAULT 'thin'::text, p_bottom text DEFAULT 'thin'::text, p_left text DEFAULT 'thin'::text, p_right text DEFAULT 'thin'::text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, OUT ret_val integer) RETURNS record
    LANGUAGE plpgsql
    AS $$
DECLARE
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
     v_borders pgexcel_generator.tp_border;
BEGIN
    ----PERFORM pgexcel_generator.Init();
    IF  COALESCE(array_length(v_workbook.borders, 1), 0)  > 0 THEN
        FOR b IN 0..COALESCE(array_length(v_workbook.borders, 1), 0) - 1 LOOP
            IF (COALESCE(v_workbook.borders[b].top, 'x') = COALESCE(p_top, 'x') AND
                COALESCE(v_workbook.borders[b].bottom, 'x') = COALESCE(p_bottom, 'x') AND
                COALESCE(v_workbook.borders[b].left, 'x') = COALESCE(p_left, 'x') AND
                COALESCE(v_workbook.borders[b].right, 'x') = COALESCE(p_right, 'x')
              ) THEN
                ret_val = b;
            END IF;
        END LOOP;
    END IF;
    t_ind := COALESCE(array_length(v_workbook.borders, 1), 0) ;
    v_borders =  v_workbook.borders[t_ind];
	v_borders.top := p_top;
    v_borders.bottom := p_bottom;
   v_borders.left := p_left;
     v_borders.right := p_right;
	  v_workbook.borders[t_ind] = v_borders;
       p_workbook  = v_workbook;
    ret_val = t_ind;
END;
$$;


--
-- Name: get_fill(text, text, pgexcel_generator.tp_book); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.get_fill(p_patterntype text, p_fgrgb text DEFAULT NULL::text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, OUT ret_val integer) RETURNS record
    LANGUAGE plpgsql
    AS $$
DECLARE
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
   v_fills pgexcel_generator.tp_fill;
BEGIN
    --PERFORM pgexcel_generator.Init();

    IF COALESCE(array_length(v_workbook.fills, 1), 0)  > 0 THEN
        FOR f IN 0..COALESCE(array_length(v_workbook.fills, 1), 0)  - 1 LOOP
            IF  v_workbook.fills[f].patternType = p_patterntype AND
              COALESCE( v_workbook.fills[f].fgRGB, 'x') = COALESCE(UPPER(p_fgrgb), 'x') THEN
                ret_val = f;
            END IF;
        END LOOP;
    END IF;
    t_ind := COALESCE(array_length(v_workbook.fills, 1), 0) ;
    v_fills = v_workbook.fills[t_ind];
    v_fills.patternType := p_patternType;
    v_fills.fgRGB := upper( p_fgRGB );
    v_workbook.fills[t_ind] = v_fills;

    p_workbook  = v_workbook;
    ret_val = t_ind;
END;
$$;


--
-- Name: get_font(text, integer, double precision, integer, boolean, boolean, boolean, text, pgexcel_generator.tp_book); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.get_font(p_name text, p_family integer DEFAULT 2, p_fontsize double precision DEFAULT 11, p_theme integer DEFAULT 1, p_underline boolean DEFAULT false, p_italic boolean DEFAULT false, p_bold boolean DEFAULT false, p_rgb text DEFAULT NULL::text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, OUT ret_val integer) RETURNS record
    LANGUAGE plpgsql
    AS $$
/* this is a hex ALPHA Red Green Blue value */
DECLARE
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book := p_workbook;
    v_fonts pgexcel_generator.tp_font;
BEGIN
    --PERFORM pgexcel_generator.Init();

    IF COALESCE(array_length(v_workbook.fonts, 1), 0)  > 0 THEN
        FOR f IN 0..COALESCE(array_length(v_workbook.fonts, 1), 0)  - 1 LOOP
            IF (v_workbook.fonts[f].name = p_name AND
                v_workbook.fonts[f].family = p_family AND
                v_workbook.fonts[f].fontsize = p_fontsize AND
                v_workbook.fonts[f].theme = p_theme AND
                v_workbook.fonts[f].underline = p_underline AND
                v_workbook.fonts[f].italic = p_italic AND
                v_workbook.fonts[f].bold = p_bold AND
                (v_workbook.fonts[f].rgb = p_rgb OR (v_workbook.fonts[f].rgb IS NULL AND p_rgb IS NULL))
              ) THEN
                ret_val = f;
            END IF;
        END LOOP;
    END IF;
    t_ind := COALESCE(array_length(v_workbook.fonts, 1), 0) ;
    v_fonts = v_workbook.fonts[t_ind];
    v_fonts.name := p_name;
     v_fonts.family := p_family;
     v_fonts.fontsize := p_fontsize;
     v_fonts.theme := p_theme;
     v_fonts.underline := p_underline;
     v_fonts.italic := p_italic;
     v_fonts.bold := p_bold;
     v_fonts.rgb := p_rgb;
     v_workbook.fonts[t_ind] = v_fonts;

     p_workbook  = v_workbook;
     ret_val = t_ind;
END;
$$;


--
-- Name: get_numfmt(text, pgexcel_generator.tp_book); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.get_numfmt(p_format text DEFAULT NULL::text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, OUT ret_val integer) RETURNS record
    LANGUAGE plpgsql
    AS $$
DECLARE
    t_cnt INTEGER;
    t_numFmtId INTEGER;
    v_workbook pgexcel_generator.tp_book := p_workbook;
    v_numFmts pgexcel_generator.tp_numFmt;
    v_numFmtIndexes numeric;
BEGIN
    --PERFORM pgexcel_generator.Init();
    IF p_format IS NULL THEN
        ret_val = 0;
    END IF;
    t_cnt :=   COALESCE(array_length(v_workbook.numFmts, 1), 0);

    FOR i IN 1..t_cnt LOOP
        IF v_workbook.numFmts[i].formatCode = p_format THEN
            t_numFmtId := v_workbook.numFmts[i].numFmtId;

            EXIT;
        END IF;
    END LOOP;
    IF t_numFmtId IS NULL THEN
        t_numFmtId :=
        CASE
            WHEN t_cnt = 0 THEN 164
            ELSE v_workbook.numFmts[t_cnt].numFmtId + 1
        END;
        t_cnt := t_cnt + 1;
        v_numFmts = v_workbook.numFmts[t_cnt];
        v_numFmts.numFmtId := t_numFmtId;
        v_numFmts.formatCode := p_format;
        v_workbook.numFmts[t_cnt] = v_numFmts;
        v_workbook.numFmtIndexes[t_numFmtId] := t_cnt::numeric;
    END IF;
      p_workbook  = v_workbook;
    ret_val = t_numFmtId;
END;
$$;


--
-- Name: get_xfid(integer, integer, integer, integer, integer, integer, integer, pgexcel_generator.tp_alignment, pgexcel_generator.tp_book); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.get_xfid(p_sheet integer, p_col integer, p_row integer, p_numfmtid integer DEFAULT NULL::integer, p_fontid integer DEFAULT NULL::integer, p_fillid integer DEFAULT NULL::integer, p_borderid integer DEFAULT NULL::integer, p_alignment pgexcel_generator.tp_alignment DEFAULT NULL::pgexcel_generator.tp_alignment, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, OUT ret_val text) RETURNS record
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;    t_cnt INTEGER;
    t_XfId INTEGER;
    t_XF pgexcel_generator.tp_xf_fmt;
    t_col_XF pgexcel_generator.tp_xf_fmt ;
    t_row_XF pgexcel_generator.tp_xf_fmt ;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
    v_alignment pgexcel_generator.tp_alignment ;
BEGIN
    --PERFORM pgexcel_generator.Init();
    IF v_workbook.sheets[p_sheet].col_fmts[p_col] THEN
        BEGIN
            t_col_XF := v_workbook.sheets[p_sheet].col_fmts[p_col];
        END;
    END IF;
    IF v_workbook.sheets[p_sheet].row_fmts[p_row] THEN
        BEGIN
          t_row_XF := v_workbook.sheets[p_sheet].row_fmts[p_row]  ;
        END;
    END IF;
      t_XF.numFmtId := coalesce( p_numFmtId, t_col_XF.numFmtId, t_row_XF.numFmtId, 0 );
      t_XF.fontId   := coalesce( p_fontId, t_col_XF.fontId, t_row_XF.fontId, 0 );
      t_XF.fillId   := coalesce( p_fillId, t_col_XF.fillId, t_row_XF.fillId, 0 );
      t_XF.borderId := coalesce( p_borderId, t_col_XF.borderId, t_row_XF.borderId, 0 );
      t_XF.alignment:= coalesce( p_alignment, t_col_XF.alignment, t_row_XF.alignment );
	  v_alignment = t_XF.alignment;
    IF (t_XF.numFmtId + t_XF.fontId + t_XF.fillId + t_XF.borderId = 0 AND
       v_alignment.vertical IS NULL AND v_alignment.horizontal IS NULL AND
       NOT COALESCE(v_alignment.wrapText, FALSE)) THEN
        ret_val = NULL;
    END IF;
    IF t_XF.numFmtId > 0 THEN
        CALL pgexcel_generator.set_col_width(p_sheet, p_col, v_workbook.numFmts[v_workbook.numFmtIndexes[t_XF.numFmtId]].formatCode,v_workbook);
    END IF;
    t_cnt :=  COALESCE(array_length(v_workbook.cellXfs, 1), 0) ;
    FOR i IN 1..t_cnt LOOP
        IF ( v_workbook.cellXfs[i].numFmtId = t_XF.numFmtId AND
            v_workbook.cellXfs[i].fontId = t_XF.fontId AND
            v_workbook.cellXfs[i].fillId = t_XF.fillId AND
            v_workbook.cellXfs[i].borderId = t_XF.borderId AND
            COALESCE( v_workbook.cellXfs[i].alignment.vertical, 'x') = COALESCE((t_XF.alignment).vertical, 'x') AND
            COALESCE( v_workbook.cellXfs[i].alignment.horizontal, 'x') = COALESCE((t_XF.alignment).horizontal, 'x') AND
            COALESCE( v_workbook.cellXfs[i].alignment.wrapText, FALSE) = COALESCE((t_XF.alignment).wrapText, FALSE)
          ) THEN
		  -- raise  notice 'inside if';
            t_XfId := i;
            EXIT;
        END IF;
    END LOOP;
    IF t_XfId IS NULL THEN
        t_cnt := t_cnt + 1;
        t_XfId := t_cnt;
        v_workbook.cellXfs[t_cnt] := t_XF;
    END IF;
    p_workbook  = v_workbook;
    --ret_val = CONCAT_WS('', 's="', t_XfId, '"');
	ret_val:= CONCAT_WS('', 's="', t_XfId, '"');
END;
$$;


--
-- Name: hyperlink(integer, integer, text, text, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.hyperlink(p_col integer, p_row integer, p_url text, p_value text DEFAULT NULL::text, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book := p_workbook;
    t_sheet INTEGER := COALESCE(p_sheet,COALESCE(array_length(v_workbook.sheets, 1), 0));
    v_rows  pgexcel_generator.tp_cells;
      v_rows_tab  pgexcel_generator.tp_cells[];
  	v_sheet pgexcel_generator.tp_sheet;
  	v_cell pgexcel_generator.tp_cell;
    v_hyperlinks pgexcel_generator.tp_hyperlink;
		v_font integer;
BEGIN
    ----PERFORM pgexcel_generator.Init();
	v_sheet = v_workbook.sheets[t_sheet];
	v_rows =  v_workbook.sheets[t_sheet].rows[p_row];
	v_cell = v_workbook.sheets[t_sheet].rows[p_row].cell[p_col];
  v_hyperlinks = v_workbook.sheets[t_sheet].hyperlinks[t_ind];
-- raise  notice 'test2';
select  * from pgexcel_generator.add_string( COALESCE( p_value, p_url ) ,v_workbook) into rec;
    v_cell.value := rec.ret_val;
		v_workbook := rec.p_workbook;
		select * from pgexcel_generator.get_font('Calibri'::TEXT, p_theme => 10, p_underline => TRUE::BOOLEAN,v_workbook) into rec;
    v_font := rec.ret_val;
		v_workbook := rec.p_workbook;
		select * from pgexcel_generator.get_xfid(t_sheet, p_col, p_row, NULL::INTEGER, v_font,v_workbook) into rec;
		v_cell.style:=CONCAT_WS('', 't="s" ',rec.ret_val );
   	v_workbook := rec.p_workbook;
    t_ind := COALESCE(array_length(v_workbook.sheets[t_sheet].hyperlinks, 1), 0) + 1;
   -- raise  notice 'test3';
   v_hyperlinks.cell := CONCAT_WS('', pgexcel_generator.alfan_col(p_col), p_row);
   -- raise  notice 'test4';
    v_hyperlinks.url := p_url;
-- raise  notice 'test5';
    v_sheet.hyperlinks[t_ind] = v_hyperlinks;
    v_rows.cell[p_col] = v_cell;
  	v_sheet.rows[p_row] = v_rows;
  	v_workbook.sheets[t_sheet] = v_sheet;
-- raise  notice 'test6';
  p_workbook  = v_workbook;
  END;
$$;


--
-- Name: list_validation(integer, integer, text, text, text, text, boolean, text, text, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.list_validation(p_sqref_col integer, p_sqref_row integer, p_defined_name text, p_style text DEFAULT 'stop'::text, p_title text DEFAULT NULL::text, p_prompt text DEFAULT NULL::text, p_show_error boolean DEFAULT false, p_error_title text DEFAULT NULL::text, p_error_txt text DEFAULT NULL::text, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
/* stop, warning, information */
BEGIN
    CALL pgexcel_generator.add_validation('list'::TEXT,
                                         CONCAT_WS('', pgexcel_generator.alfan_col(p_sqref_col), p_sqref_row),
                                         p_style => LOWER(p_style),
                                         p_formula1 => p_defined_name,
                                         p_title => p_title,
                                          p_prompt => p_prompt,
                                         p_show_error => p_show_error,
                                         p_error_title => p_error_title,
                                          p_error_txt => p_error_txt,
                                           p_sheet => p_sheet,
																				 p_workbook => p_workbook);
END;
$$;


--
-- Name: list_validation(integer, integer, integer, integer, integer, integer, text, text, text, boolean, text, text, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.list_validation(p_sqref_col integer, p_sqref_row integer, p_tl_col integer, p_tl_row integer, p_br_col integer, p_br_row integer, p_style text DEFAULT 'stop'::text, p_title text DEFAULT NULL::text, p_prompt text DEFAULT NULL::text, p_show_error boolean DEFAULT false, p_error_title text DEFAULT NULL::text, p_error_txt text DEFAULT NULL::text, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $_$
/* top left */
/* bottom right */
/* stop, warning, information */
BEGIN
    CALL pgexcel_generator.add_validation('list'::TEXT,
                                         CONCAT_WS('', pgexcel_generator.alfan_col(p_sqref_col), p_sqref_row),
                                          p_style => LOWER(p_style),
                                           p_formula1 => CONCAT_WS('', '$', pgexcel_generator.alfan_col(p_tl_col), '$', p_tl_row, ':$',pgexcel_generator.alfan_col(p_br_col), '$', p_br_row),
                                            p_title => p_title,
                                            p_prompt => p_prompt,
                                             p_show_error => p_show_error,
                                              p_error_title => p_error_title,
                                               p_error_txt => p_error_txt,
                                               p_sheet => p_sheet,
																						 p_workbook => p_workbook);
END;
$_$;


--
-- Name: mergecells(integer, integer, integer, integer, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.mergecells(p_tl_col integer, p_tl_row integer, p_br_col integer, p_br_row integer, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
/* top left */
/* bottom right */
DECLARE
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
    t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
    v_sheet pgexcel_generator.tp_sheet;
BEGIN
    ----PERFORM pgexcel_generator.Init();
    t_ind :=  COALESCE(array_length(v_workbook.sheets[t_sheet].mergecells, 1), 0) + 1;
    	v_sheet = v_workbook.sheets[t_sheet];
    v_sheet.mergecells[t_ind] := concat_ws('',pgexcel_generator.alfan_col( p_tl_col ) ,p_tl_row , ':' , pgexcel_generator.alfan_col( p_br_col ) , p_br_row);

   v_workbook.sheets[t_sheet] =v_sheet;
    p_workbook  = v_workbook;
END;
$$;


--
-- Name: new_sheet(text, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.new_sheet(p_sheetname text DEFAULT NULL::text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;
   v_workbook pgexcel_generator.tp_book  := p_workbook;
    t_nr INTEGER := COALESCE(array_length(v_workbook.sheets, 1), 0)+1;
    t_ind INTEGER;
	v_sheet pgexcel_generator.tp_sheet;
BEGIN
    --PERFORM pgexcel_generator.Init();
--raise notice 'enter new_sheet';
      	v_sheet = v_workbook.sheets[t_nr];
     v_sheet.name =  COALESCE( replace(replace(replace(replace(replace(translate(p_sheetname, 'a/\[]*:?', 'a'), '&', '&amp;'), '>', '&gt;'), '<', '&lt;'),'"','&quot;'),'''','&apos;'), CONCAT_WS('', 'Sheet', t_nr));
 v_workbook.sheets[t_nr] = v_sheet;
    IF COALESCE(array_length(v_workbook.sheets, 1), 0) = 0 THEN
        v_workbook.str_cnt := 0;
    END IF;
    IF COALESCE(array_length(v_workbook.fonts, 1), 0)  = 0 THEN
        --t_ind := pgexcel_generator.get_font('Calibri'::TEXT);
				select * from pgexcel_generator.get_font(p_name => 'Calibri'::TEXT, p_workbook => v_workbook) into rec;
				t_ind := rec.ret_Val;
				v_workbook := rec.p_workbook;
    END IF;
    IF COALESCE(array_length(v_workbook.fills, 1), 0) = 0 THEN
		select * from pgexcel_generator.get_fill('none'::TEXT,NULL::TEXT,v_workbook) into rec;
		t_ind := rec.ret_Val;
		v_workbook := rec.p_workbook;
		select * from pgexcel_generator.get_fill('gray125'::TEXT,NULL::TEXT,v_workbook) into rec;
		t_ind := rec.ret_Val;
		v_workbook := rec.p_workbook;
    --    t_ind := pgexcel_generator.get_fill('none'::TEXT);
      --  t_ind := pgexcel_generator.get_fill('gray125'::TEXT);
    END IF;
    IF COALESCE(array_length(v_workbook.borders, 1), 0)  = 0 THEN
		select * from pgexcel_generator.get_border(NULL::TEXT, NULL::TEXT, NULL::TEXT, NULL::TEXT,v_workbook) into rec;
        t_ind := rec.ret_val;
v_workbook := rec.p_workbook;
    END IF;

--raise notice 'sheet name: %', p_sheetname;
	 p_workbook  = v_workbook;
END;
$$;


--
-- Name: orafmt2excel(text); Type: FUNCTION; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE FUNCTION pgexcel_generator.orafmt2excel(p_format text DEFAULT NULL::text) RETURNS text
    LANGUAGE plpgsql
    AS $$
DECLARE
    t_format CHARACTER VARYING(1000) := substr(p_format, 1, 1000);
BEGIN
    t_format := replace(replace(t_format, 'hh24', 'hh'), 'hh12', 'hh');
    t_format := replace(t_format, 'mi', 'mm');
    t_format := replace(replace(replace(t_format, 'AM', '~~'), 'PM', '~~'), '~~', 'AM/PM');
    t_format := replace(replace(replace(t_format, 'am', '~~'), 'pm', '~~'), '~~', 'AM/PM');
    t_format := replace(replace(t_format, 'day', 'DAY'), 'DAY', 'dddd');
    t_format := replace(replace(t_format, 'dy', 'DY'), 'DAY', 'ddd');
    t_format := replace(replace(t_format, 'RR', 'RR'), 'RR', 'YY');
    t_format := replace(replace(t_format, 'month', 'MONTH'), 'MONTH', 'mmmm');
    t_format := replace(replace(t_format, 'mon', 'MON'), 'MON', 'mmm');
    RETURN t_format;
END;
$$;


--
-- Name: query2sheet(text, boolean, text, text, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.query2sheet(p_sql text, p_column_headers boolean DEFAULT true, p_directory text DEFAULT NULL::text, p_filename text DEFAULT NULL::text, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;
    t_sheet INTEGER;
    t_c NUMERIC(38);
    t_col_cnt NUMERIC(38);
    v_Val varchar;
    t_bulk_size INTEGER := 200;
    t_r NUMERIC(38);
    t_cur_row INTEGER;
   --rec record;
   rec1 record;
   v_table varchar;
   i INTEGER := 0 ;
   n INTEGER := 0 ;
   c INTEGER := 0 ;
   d INTEGER := 0 ;
    key text;
   val text;
   v_workbook pgexcel_generator.tp_book  := p_workbook;
    col_cnt integer := 1;
	row_num integer;
	add_string_var int;
BEGIN
--raise notice 'start of query2sheet';
    --PERFORM pgexcel_generator.Init();
    IF p_sheet IS NULL THEN
        CALL pgexcel_generator.new_sheet(v_workbook);
    END IF;
t_sheet := p_sheet;

execute 'drop table if exists test_dbms_sql';
execute 'drop table if exists col_types';
 execute 'create temp table if not exists test_dbms_sql as '||p_sql ;
 execute 'create temp table if not exists col_types(column_name, data_type) as select column_name, data_type from information_schema.columns WHERE table_name   = ''test_dbms_sql''';
 --raise notice 'sql query: %',p_sql;
    for rec in  execute 'select column_name from col_types'
    loop
      if p_column_headers
      then

--select pgexcel_generator.add_string(rec.column_name) into add_string_var;
        call pgexcel_generator.cell( p_col=> col_cnt, p_row => 1, p_value => rec.column_name::text,p_sheet => t_sheet ,p_workbook => v_workbook);

      end if;
	  col_cnt := col_cnt+1;
    end loop;
    t_cur_row :=
    CASE
        WHEN p_column_headers THEN 2
        ELSE 1
    END;
    t_sheet := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));

  --v_sql := concat_ws('','SELECT hstore(t) FROM (',psql,') t');
row_num = 0;
        for rec in EXECUTE p_sql
         LOOP

		 --raise notice 'rec process time: %', clock_timestamp();

		 col_cnt := 1;
         for key, val in
           select * from json_each_text(row_to_json(rec))
             loop

			if val is not null then
			--raise notice 'val ,% ',val;
			case
              when key  in (select column_name from col_types where (lower(data_type) = lower('DATE') or lower(data_type)  like lower('TIMESTAMP%'))
                           ) then
                  call pgexcel_generator.cell(p_col => col_cnt,p_row => t_cur_row + row_num, p_value => val::TIMESTAMP WITHOUT TIME ZONE, p_sheet => t_sheet ,p_workbook => v_workbook);
             when key  in (select column_name from col_types where (lower(data_type) = lower('TEXT') or lower(data_type)  like lower('%CHAR%'))
                          ) then
                   call pgexcel_generator.cell(p_col =>  col_cnt, p_row => t_cur_row + row_num, p_value => val::TEXT, p_sheet => t_sheet ,p_workbook => v_workbook);

             when key  in (select column_name from col_types where (lower(data_type) like lower('%NUM%') or lower(data_type) = lower('double Precision') or lower(data_type)  like lower('%INT%'))
                          ) then
                  call pgexcel_generator.cell(p_col =>  col_cnt, p_row => t_cur_row + row_num, p_value => val::double Precision, p_sheet => t_sheet ,p_workbook => v_workbook);

                  --end if;
				  else null;
				  end case;
				  end if;
				  col_cnt= col_cnt+1;
            end loop ;
             t_cur_row := t_cur_row + 1;
            end loop;
	p_workbook = v_workbook;

    EXCEPTION
        WHEN others THEN
            RAISE;
END;
$$;


--
-- Name: save(text, text, pgexcel_generator.tp_book, pgexcel_generator.tp_authors[]); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.save(p_directory text, p_filename text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, INOUT p_authors pgexcel_generator.tp_authors[] DEFAULT NULL::pgexcel_generator.tp_authors[])
    LANGUAGE plpgsql
    AS $$
declare
		 v_finish_txt text;
		 rec record;
BEGIN
select * from pgexcel_generator.finish(p_directory=>p_directory,p_filename => p_filename,p_workbook =>p_workbook,p_authors => p_authors) into rec;
 v_finish_txt := rec.ret_Val;
 p_workbook := rec.p_workbook;
 p_authors := rec.p_authors;
    --CALL pgexcel_generator.blob2file(v_finish_txt, p_directory, p_filename);
END;
$$;


--
-- Name: set_autofilter(integer, integer, integer, integer, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.set_autofilter(p_column_start integer DEFAULT NULL::integer, p_column_end integer DEFAULT NULL::integer, p_row_start integer DEFAULT NULL::integer, p_row_end integer DEFAULT NULL::integer, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
    t_ind INTEGER;
    v_workbook pgexcel_generator.tp_book := p_workbook;
   t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
   	v_sheet pgexcel_generator.tp_sheet;
    v_autofilters pgexcel_generator.tp_autofilter;
BEGIN
    --PERFORM pgexcel_generator.Init();

    t_ind := 1;
    v_sheet := v_workbook.sheets[t_sheet];
    v_autofilters := v_workbook.sheets[t_sheet].autofilters[t_ind];
    v_autofilters.column_start =  p_column_start;
    v_autofilters.column_end = p_column_end;
    v_autofilters.row_start = p_row_start;
    v_autofilters.row_end = p_row_end;
    v_sheet.autofilters[t_ind]  = v_autofilters;
    v_workbook.sheets[t_sheet] = v_sheet;
    CALL pgexcel_generator.defined_name(p_column_start, p_row_start, p_column_end, p_row_end, '_xlnm._FilterDatabase'::TEXT, t_sheet, t_sheet - 1,v_workbook);

    p_workbook  = v_workbook;
END;
$$;


--
-- Name: set_col_width(integer, integer, text, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.set_col_width(p_sheet integer, p_col integer, p_format text, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
rec record;    t_width DOUBLE PRECISION;
    t_nr_chr INTEGER;
    v_workbook pgexcel_generator.tp_book  := p_workbook;
    	v_sheets pgexcel_generator.tp_sheet;
BEGIN
    --PERFORM pgexcel_generator.Init();
    IF p_format IS NULL THEN
        RETURN;
    END IF;
    IF position(';' in p_format) > 0 THEN
        t_nr_chr := LENGTH(TRANSLATE(pgexcel_generator.substr(p_format, 1, position(';' in p_format) - 1), E'a\\"', 'a'));
    ELSE
        t_nr_chr := LENGTH(TRANSLATE(p_format, E'a\\"', 'a'));
    END IF;
    t_width := TRUNC(((t_nr_chr * 7 + 5) / 7 * 256)) / 256;
    /* assume default 11 point Calibri */
v_sheets = v_workbook.sheets[p_sheet];
    IF v_workbook.sheets[p_sheet].widths[p_col] is not null THEN
            v_sheets.widths[p_col] :=
                    greatest( v_sheets.widths[p_col]
                            , t_width
                            );
    ELSE
            v_sheets.widths[p_col] := greatest( t_width, 8.43 );
    END IF;
v_workbook.sheets[p_sheet] = v_sheets;

        p_workbook  = v_workbook;
END;
$$;


--
-- Name: set_column(integer, integer, integer, integer, integer, pgexcel_generator.tp_alignment, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.set_column(p_col integer, p_numfmtid integer DEFAULT NULL::integer, p_fontid integer DEFAULT NULL::integer, p_fillid integer DEFAULT NULL::integer, p_borderid integer DEFAULT NULL::integer, p_alignment pgexcel_generator.tp_alignment DEFAULT NULL::pgexcel_generator.tp_alignment, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
   v_workbook pgexcel_generator.tp_book := p_workbook;
    t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
    v_sheet pgexcel_generator.tp_sheet;
  	v_col_fmts pgexcel_generator.tp_XF_fmt;
BEGIN
    ----PERFORM pgexcel_generator.Init();
    v_sheet = v_workbook.sheets[t_sheet];
    v_col_fmts = v_workbook.sheets[t_sheet].col_fmts[p_col];
    v_col_fmts.numFmtId = p_numfmtid;
    v_col_fmts.fontId = p_fontid;
    v_col_fmts.fillId =  p_fillid;
    v_col_fmts.borderId =  p_borderid;
    v_col_fmts.alignment =  p_alignment;
    v_sheet.col_fmts[p_col] = v_col_fmts;
    v_workbook.sheets[t_sheet] = v_sheet;
      p_workbook  = v_workbook;
END;
$$;


--
-- Name: set_column_width(integer, numeric, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.set_column_width(p_col integer, p_width numeric, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
    v_workbook pgexcel_generator.tp_book := p_workbook;
		v_sheet pgexcel_generator.tp_sheet;
BEGIN
    v_sheet = v_workbook.sheets[COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0))];
    v_sheet.widths[p_col] := p_width;
	 v_workbook.sheets[COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0))]  =v_sheet ;
  p_workbook  = v_workbook;
END;
$$;


--
-- Name: set_row(integer, integer, integer, integer, integer, pgexcel_generator.tp_alignment, integer, pgexcel_generator.tp_book); Type: PROCEDURE; Schema: xlsx_builder_pkg; Owner: -
--

CREATE OR REPLACE PROCEDURE  pgexcel_generator.set_row(p_row integer, p_numfmtid integer DEFAULT NULL::integer, p_fontid integer DEFAULT NULL::integer, p_fillid integer DEFAULT NULL::integer, p_borderid integer DEFAULT NULL::integer, p_alignment pgexcel_generator.tp_alignment DEFAULT NULL::pgexcel_generator.tp_alignment, p_sheet integer DEFAULT NULL::integer, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book)
    LANGUAGE plpgsql
    AS $$
DECLARE
v_workbook pgexcel_generator.tp_book := p_workbook;
 t_sheet INTEGER := COALESCE(p_sheet, COALESCE(array_length(v_workbook.sheets, 1), 0));
 v_sheet pgexcel_generator.tp_sheet;
v_row_fmts pgexcel_generator.tp_XF_fmt;
BEGIN
    ----PERFORM pgexcel_generator.Init();
    v_sheet = v_workbook.sheets[t_sheet];
    v_row_fmts = v_workbook.sheets[t_sheet].row_fmts[p_row];
    v_row_fmts.numFmtId = p_numfmtid;
    v_row_fmts.fontId = p_fontid;
    v_row_fmts.fillId = p_fillid;
    v_row_fmts.borderId = p_borderid;
    v_row_fmts.alignment = p_alignment;
    v_sheet.row_fmts[p_row] = v_row_fmts;
    v_workbook.sheets[t_sheet] = v_sheet;
     p_workbook  = v_workbook;
END;
$$;

