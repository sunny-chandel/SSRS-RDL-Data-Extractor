#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import csv
import xml.etree.ElementTree as ET

def get_report_data(filepath):
    """
    Extracts data from a single RDL report file.
    Returns a tuple containing report name, data source name,
    dataset name, and connection string.
    """
    try:
        tree = ET.parse(filepath)
        root = tree.getroot()

        report_name = os.path.splitext(os.path.basename(filepath))[0]

        data_sources = root.findall('{http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition}DataSources')
        data_sets = root.findall('{http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition}DataSets')

        results = []
        for data_sets_child in data_sets[0].getchildren():

            data_set = data_sets_child.attrib['Name']
            for data_set_child in data_sets_child.getchildren():
                Statement = 'NA'
                for child in data_set_child:
                    # print(child)
                    if child.tag.endswith('DataSourceName') and child.text:
                        DataSourceName = child.text.strip()
                        # print((DataSourceName))

                    elif child.tag.endswith('CommandText') and child.text:
                        CommandText = child.text.strip()
                        #print(CommandText)

                    # elif child.tag.endswith('Statement') and child.text:
                    #     Statement = child.text.strip()
                    #     print(Statement)

            results.append((report_name, DataSourceName, data_set,CommandText))

        return results
    except Exception as e:
        print(f"Error occurred while processing file: {filepath}")
        print(f"Error message: {str(e)}")
        return []

def write_report_data_to_csv(report_data, output_path):
    """
    Writes report data to a CSV file at the specified output path.
    """
    with open(output_path, mode='w', newline='') as csv_file:
        fieldnames = ['Report Name', 'DataSource Name', 'Dataset','DAX_query','Statement']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()

        for data in report_data:
            writer.writerow({
                'Report Name': data[0],
                'DataSource Name': data[1],
                'Dataset': data[2],
                'DAX_query':data[3],
                # 'Statement':data[4]
            })

def extract_report_data(folder_path):
    """
    Extracts report data from all RDL files in the specified folder.
    Returns a list of tuples containing report name, data source name,
    dataset name, and connection string.
    """
    report_data = []

    for filename in os.listdir(folder_path):
        if filename.endswith('.rdl'):
            filepath = os.path.join(folder_path, filename)
            report_data += get_report_data(filepath)

    return report_data

def main():
    """ac
    Main function that calls other functions to extract report data
    and write it to a CSV file.
    """
    folder_path = r"C:\Users\mkum3055\Downloads\SSRS Reports\SSRS RDL Files"
    output_path = '104.csv'
    report_data = extract_report_data(folder_path)
    write_report_data_to_csv(report_data, output_path)

if __name__ == "__main__":
    main()

