#!/usr/bin/env python
# coding: utf-8

# In[21]:


import xml.etree.ElementTree as ET

# Load the XML file
tree = ET.parse(r"C:\Users\sunny.chandel\Desktop\RDL_27\POSPSA APSD YTD Day.rdl")
root = tree.getroot()

# Define the namespace
namespace = '{http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition}'

# Find the <DataSets> element using the namespace
datasets_element = root.find(f'{namespace}DataSets')

# Extract the <Code> element from the root
code_element = root.find(f'{namespace}Code')

# Print or process the information from the <Code> element
if code_element is not None:
    code_content = code_element.text
    print("<Code> Content:")
    print(code_content)
    print("-" * 30)

# Iterate through each <Dataset> element within <DataSets>
for dataset_element in datasets_element:
    dataset_name = dataset_element.get('Name')  # Extract dataset name from the 'Name' attribute
    
    data_source_element = dataset_element.find(f'{namespace}Query/{namespace}DataSourceName')
    data_source = data_source_element.text if data_source_element is not None else None
    
    query_element = dataset_element.find(f'{namespace}Query/{namespace}CommandText')
    query = query_element.text if query_element is not None else None
    
    fields_element = dataset_element.find(f'{namespace}Fields')
    
    # Gather information from each <Field> element within <Fields> for the current dataset
    fields_info = []
    for field_element in fields_element:
        field_name = field_element.get('Name')
        data_field_element = field_element.find(f'{namespace}DataField')
        data_field = data_field_element.text if data_field_element is not None else None
        user_defined_element = field_element.find(f'{namespace}UserDefined')
        user_defined = user_defined_element.text if user_defined_element is not None else None
        
        fields_info.append({
            'Field Name': field_name,
            'Data Field': data_field,
            'User Defined': user_defined
        })
    
    # Print information for the current dataset
    print("Dataset Name:", dataset_name)
    print("Data Source:", data_source)
    print("Query:", query)
    for field in fields_info:
        print("Field Name:", field['Field Name'])
        print("Data Field:", field['Data Field'])
        print("User Defined:", field['User Defined'])
    print("-" * 30)  # Print separator after completion of a dataset


# In[62]:


import xml.etree.ElementTree as ET
import pandas as pd

# Define the directory path containing RDL files
rdl_directory = r"C:\Users\sunny.chandel\Desktop\RDL_27"

# Define the namespace
namespace = '{http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition}'

# Create a list to hold exception information
exceptions_info = []

# Initialize lists for datasets and parameters information
datasets_info = []
parameters_info = []

# Iterate through each RDL file in the directory
for filename in os.listdir(rdl_directory):
    if filename.endswith(".rdl"):
        # Load the XML file
        rdl_path = os.path.join(rdl_directory, filename)
        try:
            tree = ET.parse(rdl_path)
            root = tree.getroot()

            # Extract report name from the file name (removing .rdl extension)
            report_name = os.path.splitext(filename)[0]
            # Extract the <Code> element from the root
            code_element = root.find(f'{namespace}Code')
            code_text = code_element.text.strip() if code_element is not None else None

            # Find the <DataSets> element using the namespace
            datasets_element = root.find(f'{namespace}DataSets')

            # Iterate through each <Dataset> element within <DataSets>
            for dataset_element in datasets_element:
                dataset_name = dataset_element.get('Name')  # Extract dataset name from the 'Name' attribute

                data_source_element = dataset_element.find(f'{namespace}Query/{namespace}DataSourceName')
                data_source = data_source_element.text if data_source_element is not None else None

                query_element = dataset_element.find(f'{namespace}Query/{namespace}CommandText')
                query = query_element.text if query_element is not None else None

                if query and query.startswith('='):
                    query = "'" + query

                fields_element = dataset_element.find(f'{namespace}Fields')

                # Gather information from each <Field> element within <Fields> for the current dataset
                for field_element in fields_element:
                    field_name = field_element.get('Name')
                    data_field_element = field_element.find(f'{namespace}DataField')
                    data_field = data_field_element.text if data_field_element is not None else None

                    datasets_info.append({
                        'Report Name': report_name,
                        'Dataset Name': dataset_name,
                        'Data Source': data_source,
                        'Query': query,
                        'Field Name': field_name,
                        'Data Field': data_field,
                        'Code': code_text
                    })

            # Find the <ReportParameters> element using the namespace
            parameters_element = root.find(f'{namespace}ReportParameters')
            if parameters_element is not None:
                # Iterate through each <ReportParameter> element within <ReportParameters>
                for parameter_element in parameters_element:
                    try:
                        parameter_name = parameter_element.get('Name')

                        parameter_data_type_element = parameter_element.find(f'{namespace}DataType')
                        if parameter_data_type_element is not None:
                            parameter_data_type = parameter_data_type_element.text
                        else:
                            parameter_data_type = None

                        parameter_default_value_element = parameter_element.find(f'{namespace}DefaultValue/{namespace}Values/{namespace}Value')
                        if parameter_default_value_element is not None:
                            parameter_default_value = parameter_default_value_element.text
                        else:
                            parameter_default_value = None

                        parameter_prompt_element = parameter_element.find(f'{namespace}Prompt')
                        if parameter_prompt_element is not None:
                            parameter_prompt = parameter_prompt_element.text
                        else:
                            parameter_prompt = None

                        valid_values_element = parameter_element.find(f'{namespace}ValidValues')
                        if valid_values_element is not None:
                            valid_values = []
                            for value_element in valid_values_element:
                                if value_element.tag == f'{namespace}ParameterValues':
                                    for sub_value_element in value_element:
                                        value = sub_value_element.find(f'{namespace}Value').text
                                        label = sub_value_element.find(f'{namespace}Label').text
                                        valid_values.append({'Value': value, 'Label': label})
                                elif value_element.tag == f'{namespace}DataSetReference':
                                    dataset_name = value_element.find(f'{namespace}DataSetName').text
                                    value_field = value_element.find(f'{namespace}ValueField').text
                                    label_field = value_element.find(f'{namespace}LabelField').text
                                    # Handle DataSetReference case accordingly
                                elif value_element.tag == f'{namespace}Query':
                                    # Handle Query case
                                    pass
                                elif value_element.tag == f'{namespace}StaticValues':
                                    # Handle StaticValues case
                                    pass
                                elif value_element.tag == f'{namespace}DynamicValues':
                                    # Handle DynamicValues case
                                    pass
                                else:
                                    # Handle other cases as needed
                                    pass

                            parameters_info.append({
                                'Report Name': report_name,
                                'Parameter Name': parameter_name,
                                'Data Type': parameter_data_type,
                                'Default Value': parameter_default_value,
                                'Prompt': parameter_prompt,
                                'Valid Values': valid_values
                            })
                        else:
                            parameters_info.append({
                                'Report Name': report_name,
                                'Parameter Name': parameter_name,
                                'Data Type': parameter_data_type,
                                'Default Value': parameter_default_value,
                                'Prompt': parameter_prompt,
                                'Valid Values': None
                            })

                    except Exception as e:
                        print(e)  # Print exception details for debugging
                        exceptions_info.append({
                            'Report Name': report_name,
                            'Parameter Name': parameter_name,
                            'Exception Details': str(e)
                        })
                else:
                    parameters_info.append({
                                'Report Name': report_name,
                                'Parameter Name': None,
                                'Data Type': None,
                                'Default Value': None,
                                'Prompt': None,
                                'Valid Values': None
                            })
                    
                        

        except Exception as e:
            exceptions_info.append({
                'File Name': filename,
                'Exception Details': str(e)
            })

# Create DataFrames from the lists
exceptions_df = pd.DataFrame(exceptions_info)
df_datasets = pd.DataFrame(datasets_info)
df_parameters = pd.DataFrame(parameters_info)

# Define the Excel writer
excel_writer = pd.ExcelWriter(r"C:\Users\sunny.chandel\Desktop\RDL_27\output1.xlsx", engine='xlsxwriter')

# Write the DataFrames to separate sheets
df_datasets.to_excel(excel_writer, sheet_name='Datasets', index=False)
df_parameters.to_excel(excel_writer, sheet_name='Parameters', index=False)
exceptions_df.to_excel(excel_writer, sheet_name='Exceptions', index=False)

# Save the Excel file
excel_writer.save()
print('done!!!')


# In[ ]:




