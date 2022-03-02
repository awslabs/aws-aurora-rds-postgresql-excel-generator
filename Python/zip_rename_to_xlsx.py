import json
import os, shutil
import boto3
import botocore.exceptions
from zipfile import ZipFile, ZIP_DEFLATED


def lambda_handler(event, context):
    
    try:
        s3_resource = boto3.resource('s3')
        s3_client = boto3.client('s3')
        
        report_name = event.get('key1')
        report_folder = report_name+'/'
        bucket = os.environ['BUCKET_NAME']
        base_folder = os.environ['BASE_FOLDER'] + '/'
        
        print('Generating Report ' + report_name)
        files_to_zip = GetFilesToZip(s3_client, bucket, base_folder, report_folder)
        CreateLocalDirectories(base_folder, report_folder)
        DownloadFilesFromS3(s3_resource, bucket, base_folder, report_folder, files_to_zip)
        CreateSpreadsheetLocally(report_name, bucket, base_folder, report_folder, files_to_zip)
        UploadSpreadsheetToS3(report_name, s3_resource, bucket, base_folder)
        DeleteReportFilesFromS3(s3_client, bucket, base_folder, report_folder)
        print('Generated Report: ' + report_name)
        
    except Exception as e:
        print('Error generating report ', report_name)
        print('> report_name: ', report_name)
        print('> bucket: ', bucket)
        print('> Exception: ', e)
        
    finally:
        RemoveLocalDirectories(base_folder)
    
    
###
def GetFilesToZip(s3_client, bucket, base_folder, report_folder):
    try: 
        files_to_zip = []
        response = s3_client.list_objects_v2(Bucket=bucket, Prefix=base_folder+report_folder)
        all_contents = response['Contents']
        for i in all_contents:
            files_to_zip.append(str(i['Key']))
        return files_to_zip
    except Exception as e:
        print('... error on function GetFilesToZip')
        raise


###
def CreateLocalDirectories(base_folder, report_folder):
    try:
        os.chdir('/tmp/')
        os.mkdir(base_folder)
        os.mkdir(base_folder+report_folder)
        os.mkdir(base_folder+report_folder+'_rels/')
        os.mkdir(base_folder+report_folder+'docProps/')
        os.mkdir(base_folder+report_folder+'xl/')
        os.mkdir(base_folder+report_folder+'xl/_rels/')
        os.mkdir(base_folder+report_folder+'xl/theme/')
        os.mkdir(base_folder+report_folder+'xl/worksheets/')
    except Exception as e:
        print('... error on function CreateLocalDirectories')
        raise
    

###
def DownloadFilesFromS3(s3_resource, bucket, base_folder, report_folder, files_to_zip):
    for file_name in files_to_zip:
        try:
            s3_resource.Bucket(bucket).download_file(file_name, file_name)
        except botocore.exceptions.ClientError as e:
            print('... error on function DownloadFilesFromS3')
            raise
        except Exception as e:
            print('... error on function DownloadFilesFromS3')
            print('> bucket: ' + bucket)
            print('> file_name: ' + file_name)
            raise        
    
###
def CreateSpreadsheetLocally(p_report_name, bucket, base_folder, report_folder, files_to_zip):
    try:
        report_folder = p_report_name+'/'
        report_file_name = base_folder + p_report_name + '.xlsx'
        with ZipFile('/tmp/' + report_file_name, 'w', compression=ZIP_DEFLATED, allowZip64=True) as zip:    
            for file in files_to_zip:
                zip.write(file, file) #is this correct?
                zip.write(file, file.replace(base_folder+report_folder, '/') )
    except Exception as e:
        print('... error on function CreateSpreadsheetLocally')
        raise


###
def UploadSpreadsheetToS3(p_report_name, s3_resource, bucket, base_folder):
    try:
        report_file_name = base_folder + p_report_name + '.xlsx'
        s3_resource.Bucket(bucket).upload_file('/tmp/' + report_file_name, base_folder + p_report_name + '.xlsx')
    except Exception as e:
        print('... error on function UploadSpreadsheetToS3')
        raise
    
    
###
def RemoveLocalDirectories(base_folder):
    try:
        os.chdir('/tmp/')
        shutil.rmtree('/tmp/'+base_folder)
    except Exception as e:
        print('... error on function RemoveLocalDirectories')
        raise
    

###
def DeleteReportFilesFromS3(s3_client, bucket, base_folder, report_folder):
    try:
        response = s3_client.list_objects_v2(Bucket=bucket, Prefix=base_folder+report_folder)
        all_contents = response['Contents']
        keys_to_delete = {'Objects' : []}
        keys_to_delete['Objects'] = [{'Key' : k} for k in [obj['Key'] for obj in all_contents ]]
        s3_client.delete_objects(Bucket=bucket, Delete=keys_to_delete)
    except Exception as e:
        print('... error on function DeleteReportFilesFromS3')
        raise
    

###
def print_local_directories():
    main_root = os.path.join(os.getcwd(), '/tmp/')
    for path, subdirs, files in os.walk(main_root):
        print(path)
        for name in files:
            print(path + '/' + name)
    
    
