import os

def create_job_list(order_num, filepath):
    """For each line in job_entry.csv, this function will go through all of the
    .xlsx and .xlsm files and create a list of tuples containing the original
    full filepath and the filepath/name of the new file to be added."""
    for root, dirs, files in os.walk(filepath):
        # root and dirs are unused here. There may be a proper way to
        # throw away these variables.
        job_files = []
        files = [x for x in files if x.endswith(".xlsx") or x.endswith(".xlsm")]
        # Create an empty list to fill with renamed Excel docs
        # to scroll through later
        for line, file in enumerate(files,1):
            input_file = filepath+'\\'+file
            output_file = filepath+'\\'+'0'+str(order_num)+'-'+str(line)+'-1 Job Entry.xlsx'
            job_files.append((input_file,output_file))

        return(job_files)
        # Send list of job files to be scanned for materials & operations
