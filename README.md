
A PL/PGSQL Package to create EXCEL workbooks. This package enables developers to export data from an RDS or Aurora PostgreSQL database to Excel files using simple PL/PGSQL functions and procedures.


1. Create aws_s3 extension
```
excel_test=> create extension aws_s3 CASCADE;
NOTICE:  installing required extension "aws_commons"
CREATE EXTENSION
> excel_test=> \dx
                           List of installed extensions
    Name     | Version |   Schema   |                 Description
-------------+---------+------------+---------------------------------------------
 aws_commons | 1.2     | public     | Common data types across AWS services
 aws_s3      | 1.1     | public     | AWS S3 extension for importing data from S3
 plpgsql     | 1.0     | pg_catalog | PL/pgSQL procedural language
(3 rows)

> excel_test=> \dx aws_*
                         List of installed extensions
    Name     | Version | Schema |                 Description
-------------+---------+--------+---------------------------------------------
 aws_commons | 1.2     | public | Common data types across AWS services
 aws_s3      | 1.1     | public | AWS S3 extension for importing data from S3
(2 rows)

```

1. Install UTL_FILE utility by running sql/utl_file_utility.sql script in the RDS/Aurora PostgreSQL database.

Note that the region is in init() function of the script. If you want different region, you can change in the function before running.
```
psql -h pg-instance.xxxx.us-east-1.rds.amazonaws.com -p 5432 -U postgres -d excel_test -f sql/utl_file_utility.sql > utl.log 2>&1 &

excel_test=> \dn utl_file_utility
       List of schemas
       Name       |  Owner
------------------+----------
 utl_file_utility | postgres
(1 row)

excel_test=> \df utl_file_utility.*
                                                                                               List of functions
      Schema      |    Name     |      Result data type      |                                                            Argument data types                                                            | Type
------------------+-------------+----------------------------+-------------------------------------------------------------------------------------------------------------------------------------------+------
 utl_file_utility | fclose      | numeric                    | p_path character varying, p_file_name character varying                                                                                   | func
 utl_file_utility | fclose_csv  | numeric                    | p_path character varying, p_file_name character varying                                                                                   | func
 utl_file_utility | fclose_xlsx | numeric                    | p_path character varying, p_file_name character varying                                                                                   | func
 utl_file_utility | fgetattr    | bigint                     | p_path character varying, p_file_name character varying                                                                                   | func
 utl_file_utility | fopen       | utl_file_utility.file_type | p_path character varying, p_file_name character varying, p_mode character DEFAULT 'W'::bpchar, OUT p_file_type utl_file_utility.file_type | func
 utl_file_utility | get_line    | text                       | p_path character varying, p_file_name character varying, p_buffer text                                                                    | func
 utl_file_utility | init        | void                       |                                                                                                                                           | func
 utl_file_utility | is_open     | boolean                    | p_path character varying, p_file_name character varying                                                                                   | func
 utl_file_utility | put_line    | boolean                    | p_path character varying, p_file_name character varying, p_line text, p_flag character DEFAULT 'W'::bpchar                                | func
(9 rows)

```

2. Create a S3 bucket and a base folder to store the excel reports.

3. Update the bucket and base folder details by inserting a record in utl_file_utility.all_directories table
```
INSERT INTO utl_file_utility.all_directories VALUES (nextval('utl_file_utility.all_directories_id_seq'), 'REPORTS', 'pg-reports-testing', 'excel_reports');


excel_test=> INSERT INTO utl_file_utility.all_directories VALUES (nextval('utl_file_utility.all_directories_id_seq'), 'REPORTS', 'pg-reports-testing', 'excel_reports');
INSERT 0 1
excel_test=>
excel_test=>
excel_test=> SELECT * FROM utl_file_utility.all_directories;
 id | directory_name |     s3_bucket      |    s3_path
----+----------------+--------------------+---------------
  1 | REPORTS        | pg-reports-testing | excel_reports
(1 row)

excel_test=>

```
Note that REPORTS is a directory name which you will use as an input while generating reports.


3. Setting up access to an Amazon S3 bucket from the RDS or Aurora PostgreSQL instance using: https://docs.aws.amazon.com/AmazonRDS/latest/UserGuide/USER_PostgreSQL.S3Import.html#USER_PostgreSQL.S3Import.AccessPermission

5. Create Lambda - zip_rename_to_xlsx.py using steps: https://docs.aws.amazon.com/lambda/latest/dg/getting-started-create-function.html#gettingstarted-zip-function

You will have to add the following environment variables for your Lambda function.

```
BUCKET_NAME = <your bucket name>
BASE_FOLDER = <your base folder>
```

4. Giving database access to Lambda: https://docs.aws.amazon.com/AmazonRDS/latest/AuroraUserGuide/PostgreSQL-Lambda.html#PostgreSQL-Lambda-access

Following are the higher level steps:

a. Create a Policy
b. Create a Role for RDS service
c. Attach the policy to the role
d. Attach the role to the RDS/Aurora instance


5. Install Lambda extension: 

https://docs.aws.amazon.com/AmazonRDS/latest/AuroraUserGuide/PostgreSQL-Lambda.html#PostgreSQL-Lambda-overview

```
CREATE EXTENSION IF NOT EXISTS aws_lambda CASCADE;

excel_test=> CREATE EXTENSION IF NOT EXISTS aws_lambda CASCADE;
CREATE EXTENSION
excel_test=> \dx aws_la*
              List of installed extensions
    Name    | Version | Schema |      Description
------------+---------+--------+------------------------
 aws_lambda | 1.0     | public | AWS Lambda integration
(1 row)

excel_test=>
```

6. Install PGEXCEL_BUILDER_PKG by running sql/PostgreSQL_excel_generator_pkg.sql script in the RDS/Aurora PostgreSQL database

```
$ psql -h pg-instance.xxxx.us-east-1.rds.amazonaws.com -p 5432 -U postgres -d excel_test -f sql/PostgreSQL_excel_generator_pkg.sql > excel_procs.log 2>&1 &

excel_test=> \dn pgexcel*
       List of schemas
       Name        |  Owner
-------------------+----------
 pgexcel_generator | postgres
(1 row)
```

7. Create sample procedure to generate an excel report using sql/sample_proc_to_format_cells.sql script.

```
$ psql -h pg-instance.xxxx.us-east-1.rds.amazonaws.com -p 5432 -U postgres -d excel_test -f sql/sample_proc_to_format_cells.sql

excel_test=> \df sample_proc_to_format_cells
List of functions
-[ RECORD 1 ]-------+---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Schema              | public
Name                | sample_proc_to_format_cells
Result data type    |
Argument data types | p_directory character varying, p_report_name character varying, p_lambdafunctionname character varying, INOUT p_workbook pgexcel_generator.tp_book DEFAULT NULL::pgexcel_generator.tp_book, INOUT preturnval character varying DEFAULT NULL::character varying
Type                | proc
```

8. Create a sample table to export data to workbook sheet:

```
create table test_excel(id int,name varchar, eid varchar, mid varchar, nid int);

insert into test_excel values(generate_series(1,100),'name'||generate_series(1,100),'eid'||generate_series(1,100),'mid'||generate_series(1,100),generate_series(1,100));
```

9. Run sample_proc_to_format_cells procedure.

```
call sample_proc_to_format_cells('REPORTS','excel_generate_test.xlsx','arn:aws:lambda:us-east-1:xxxxx:function:CreateExcelReport1');
```

Note that you will have to use your Lambda function ARN as an input.

10. Verify excel report generated in the S3 bucket.

