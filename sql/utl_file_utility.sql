
--
-- Name: utl_file_utility; Type: SCHEMA; Schema: -; Owner: -
--

CREATE SCHEMA utl_file_utility;


--
-- Name: file_type; Type: TYPE; Schema: utl_file_utility; Owner: -
--

CREATE TYPE utl_file_utility.file_type AS (
	p_path character varying,
	p_file_name character varying
);


--
-- Name: fclose(character varying, character varying); Type: FUNCTION; Schema: utl_file_utility; Owner: -
--

CREATE FUNCTION utl_file_utility.fclose(p_path character varying, p_file_name character varying) RETURNS numeric
    LANGUAGE plpgsql
    AS $$

  DECLARE

    v_filewithpath CHARACTER varying;

    v_bucket CHARACTER varying;

    v_region CHARACTER varying;

    v_tabname CHARACTER varying;

    v_sql text;

    l_file_length NUMERIC;

    v_s3_path     VARCHAR;

        v_temp_count numeric;

  BEGIN

    perform utl_file_utility.init();

    v_region := current_setting( format( '%s.%s', 'UTL_FILE_UTILITY', 'region' ) );

    SELECT s3_bucket,

           s3_path

    INTO   v_bucket,

           v_s3_path

    FROM   utl_file_utility.all_directories

    WHERE  directory_name = p_path;

    v_tabname := concat_ws('','"',substring(p_file_name, 1,

    CASE

    WHEN strpos(p_file_name, '.') = 0 THEN

      length(p_file_name)

    ELSE

      strpos(p_file_name, '.') - 1

    END ),'"');

    v_filewithpath :=

    CASE

    WHEN nullif(v_s3_path, '') IS NULL THEN

      p_file_name

    ELSE

      concat_ws('/', v_s3_path, p_file_name)

    END ;

 

    v_sql := concat_ws('', 'select count(1) from ', v_tabname, '');

        EXECUTE v_sql into v_temp_count;

        IF v_temp_count = 0 THEN

        SELECT bytes_uploaded

    INTO   l_file_length

    FROM   aws_s3.query_export_to_s3 ('select ''''', aws_commons.create_s3_uri(v_bucket, v_filewithpath, v_region) );

        ELSE

        SELECT bytes_uploaded

    INTO   l_file_length

    FROM   aws_s3.query_export_to_s3 (concat_ws('', 'select * from ', v_tabname, ' order by ctid asc'), aws_commons.create_s3_uri(v_bucket, v_filewithpath, v_region) );

        END IF;

    v_sql := 'drop table '

    || v_tabname;

    EXECUTE v_sql;

    UPDATE utl_file_utility.file_status

    SET    status = 'Closed'

    WHERE  file_name = p_file_name

    AND    file_path = p_path;

    RETURN l_file_length;

  EXCEPTION

  WHEN OTHERS THEN

    RAISE notice 'error fclose %', SQLERRM;

    RAISE;

  END;

  $$;



--
-- Name: fopen(character varying, character varying, character); Type: FUNCTION; Schema: utl_file_utility; Owner: -
--

CREATE FUNCTION utl_file_utility.fopen(p_path character varying, p_file_name character varying, p_mode character DEFAULT 'W'::bpchar, OUT p_file_type utl_file_utility.file_type) RETURNS utl_file_utility.file_type
    LANGUAGE plpgsql
    AS $$

  DECLARE

    v_sql CHARACTER varying;

    v_cnt_stat INTEGER;

    v_cnt      INTEGER;

    v_tabname CHARACTER varying;

    v_filewithpath CHARACTER varying;

    v_region CHARACTER varying;

    v_bucket CHARACTER varying;

    v_s3_path CHARACTER varying;

  BEGIN

    /*initialize common variable */

    perform utl_file_utility.init();

    v_region := current_setting( format( '%s.%s', 'UTL_FILE_UTILITY', 'region' ) );

    SELECT s3_bucket,

           s3_path

    INTO   v_bucket,

           v_s3_path

    FROM   utl_file_utility.all_directories

    WHERE  upper(directory_name) = upper(p_path);

    

    /* set tabname*/

    v_tabname := concat_ws('','"',substring(p_file_name, 1,

    CASE

    WHEN strpos(p_file_name, '.') = 0 THEN

      length(p_file_name)

    ELSE

      strpos(p_file_name, '.') - 1

    END ),'"');

    --RAISE notice 'v_tabname ,%',v_tabname;

    v_filewithpath :=

    CASE

    WHEN nullif(v_s3_path, '') IS NULL THEN

      p_file_name

    ELSE

      concat_ws('/', v_s3_path, p_file_name)

    END ;

     --raise notice 'v_bucket %, v_filewithpath % , v_region %',v_bucket,v_filewithpath,v_region;

    /* APPEND MODE HANDLING; RETURN EXISTING FILE DETAILS IF PRESENT ELSE CREATE AN EMPTY FILE */

    IF upper(p_mode) IN ('WB',

                         'RB',

                         'AB') THEN

      v_sql := concat_ws('', 'create temp table if not exists ', v_tabname, ' (col1 bytea)');

      EXECUTE v_sql;

    ELSEIF upper(p_mode) IN ('R','A','W') THEN

      v_sql := concat_ws('', 'create temp table if not exists ', v_tabname, ' (col1 text)');

	  --raise notice 'created temp table: %', v_tabname;

      EXECUTE v_sql;

    END IF;

    -- raise notice 'v_sql %',v_sql;

    IF p_mode IN ( 'A',

                  'AB') THEN

      BEGIN

        perform aws_s3.table_import_from_s3 ( v_tabname, '', 'DELIMITER AS ''^''', aws_commons.create_s3_uri ( v_bucket, v_filewithpath , v_region) );

      EXCEPTION

      WHEN OTHERS THEN

        RAISE notice 'File load issue ,%', SQLERRM;

        IF SQLERRM NOT LIKE '%file does not exist%' THEN

          RAISE;

        END IF;

      END;

      EXECUTE concat_ws('', 'select count(*) from ', v_tabname) INTO v_cnt;

      IF v_cnt > 0 THEN

        p_file_type.p_path := v_filewithpath;

        p_file_type.p_file_name := p_file_name;

      ELSE

        perform aws_s3.query_export_to_s3('select ''''', aws_commons.create_s3_uri(v_bucket, v_filewithpath, v_region) );

        p_file_type.p_path := v_filewithpath;

        p_file_type.p_file_name := p_file_name;

      END IF;

    elseif p_mode IN ('WB',

                      'W') THEN

      perform aws_s3.query_export_to_s3('select '' ''', aws_commons.create_s3_uri(v_bucket, v_filewithpath, v_region) );

      p_file_type.p_path := v_filewithpath;

      p_file_type.p_file_name := p_file_name;

    elseif p_mode IN ('RB',

                      'R') THEN

      BEGIN

        perform aws_s3.table_import_from_s3 ( v_tabname, '', 'DELIMITER AS ''^''', aws_commons.create_s3_uri ( v_bucket, v_filewithpath , v_region) );

      EXCEPTION

      WHEN OTHERS THEN

        RAISE notice 'File not found for read mode ,%', SQLERRM;

        RAISE;

      END;

    END IF;

    SELECT count(*)

    INTO   v_cnt

    FROM   utl_file_utility.file_status

    WHERE  file_name = p_file_name

    AND    file_path = p_path;

    

    IF v_cnt = 0 THEN

      INSERT INTO utl_file_utility.file_status

                  (

                              file_name,

                              status,

                              file_path,

                              pid

                  )

                  VALUES

                  (

                              p_file_name,

                              'open',

                              p_path,

                              pg_backend_pid()

                  );

    

    ELSE

      UPDATE utl_file_utility.file_status

      SET    status = 'open' ,

             pid = pg_backend_pid()

      WHERE  file_name = p_file_name

      AND    file_path = p_path;

    

    END IF;

  EXCEPTION

  WHEN OTHERS THEN

    p_file_type.p_path := v_filewithpath;

    p_file_type.p_file_name := p_file_name;

    RAISE notice 'fopenerror,%', SQLERRM;

    RAISE;

  END;

  $$;


--
-- Name: get_line(character varying, character varying, text); Type: FUNCTION; Schema: utl_file_utility; Owner: -
--

CREATE FUNCTION utl_file_utility.get_line(p_path character varying, p_file_name character varying, p_buffer text) RETURNS text
    LANGUAGE plpgsql
    AS $$ declare v_tabname varchar;

v_sql varchar;

p_buffer text;

begin v_tabname := substring(p_file_name, 1, case when strpos(p_file_name, '.') = 0 then length(p_file_name) else strpos(p_file_name, '.') - 1 end );

v_sql := 'select string_agg(col1,E''\n'') from

' || v_tabname;

execute v_sql

into

	p_buffer;

return p_buffer;

exception

when others then raise notice 'error get_line %',

sqlerrm;

select

	*

from

	utl_file_utility.fclose(p_path ,

	p_file_name);

raise;

end;

$$;


--
-- Name: init(); Type: FUNCTION; Schema: utl_file_utility; Owner: -
--

CREATE FUNCTION utl_file_utility.init() RETURNS void
    LANGUAGE plpgsql
    AS $$

BEGIN

	  perform set_config

      ( format( '%s.%s','UTL_FILE_UTILITY', 'region' )

      , 'us-east-1'::text

      , false );

END;

$$;


--
-- Name: is_open(character varying, character varying); Type: FUNCTION; Schema: utl_file_utility; Owner: -
--

CREATE FUNCTION utl_file_utility.is_open(p_path character varying, p_file_name character varying) RETURNS boolean
    LANGUAGE plpgsql
    AS $$ declare v_cnt int;

begin

select

	count(*)

into

	v_cnt

from

	utl_file_utility.file_status

where

	file_name = p_file_name

	and file_path = p_path

	and upper(status) = 'OPEN'

	and pid in (

	select

		pid

	from

		pg_stat_activity);

if v_cnt > 0 then return true;

else return false;

end if;

exception

when others then raise notice 'error is_open %',

sqlerrm;

raise;

end;

$$;


--
-- Name: put_line(character varying, character varying, text, character); Type: FUNCTION; Schema: utl_file_utility; Owner: -
--

CREATE FUNCTION utl_file_utility.put_line(p_path character varying, p_file_name character varying, p_line text, p_flag character DEFAULT 'W'::bpchar) RETURNS boolean
    LANGUAGE plpgsql
    AS $$

/*************************************************************************** Write line **************************************************************************/

declare v_sql varchar;

v_ins_sql varchar;

v_cnt INTEGER;

v_filewithpath character varying;

v_tabname character varying;

v_bucket character varying;

v_region character varying;

begin perform utl_file_utility.init();

/* check if temp table already exist */

--v_tabname := substring(p_file_name, 1, case when strpos(p_file_name, '.') = 0 then length(p_file_name) else strpos(p_file_name, '.') - 1 end );

v_tabname := concat_ws('','"',substring(p_file_name, 1, case when strpos(p_file_name, '.') = 0 then length(p_file_name) else strpos(p_file_name, '.') - 1 end ),'"');

/* INSERT INTO TEMP TABLE */

v_ins_sql := concat_ws('', 'insert into ', v_tabname, '

values(''', p_line, ''')');

execute v_ins_sql;

return true;

exception

when others then raise notice 'Error Message : %',

sqlerrm;

select

	*

from

	utl_file_utility.fclose(p_path ,

	p_file_name);

raise;

end;

$$;


SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: all_directories; Type: TABLE; Schema: utl_file_utility; Owner: -
--

CREATE TABLE utl_file_utility.all_directories (
    id integer NOT NULL,
    directory_name character varying,
    s3_bucket character varying,
    s3_path character varying
);


--
-- Name: all_directories_id_seq; Type: SEQUENCE; Schema: utl_file_utility; Owner: -
--

CREATE SEQUENCE utl_file_utility.all_directories_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: all_directories_id_seq; Type: SEQUENCE OWNED BY; Schema: utl_file_utility; Owner: -
--

ALTER SEQUENCE utl_file_utility.all_directories_id_seq OWNED BY utl_file_utility.all_directories.id;


--
-- Name: file_status; Type: TABLE; Schema: utl_file_utility; Owner: -
--

CREATE TABLE utl_file_utility.file_status (
    id integer NOT NULL,
    file_name character varying,
    status character varying,
    pid numeric,
    file_path character varying
);


--
-- Name: file_status_id_seq; Type: SEQUENCE; Schema: utl_file_utility; Owner: -
--

CREATE SEQUENCE utl_file_utility.file_status_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: file_status_id_seq; Type: SEQUENCE OWNED BY; Schema: utl_file_utility; Owner: -
--

ALTER SEQUENCE utl_file_utility.file_status_id_seq OWNED BY utl_file_utility.file_status.id;



--
-- Name: all_directories id; Type: DEFAULT; Schema: utl_file_utility; Owner: -
--

ALTER TABLE ONLY utl_file_utility.all_directories ALTER COLUMN id SET DEFAULT nextval('utl_file_utility.all_directories_id_seq'::regclass);


--
-- Name: file_status id; Type: DEFAULT; Schema: utl_file_utility; Owner: -
--

ALTER TABLE ONLY utl_file_utility.file_status ALTER COLUMN id SET DEFAULT nextval('utl_file_utility.file_status_id_seq'::regclass);


