import os
import pandas as pd
import math
import streamlit as st
import shutil
import glob

def clear_folder(folder_path):
    files = glob.glob(os.path.join(folder_path, "*.csv"))
    for f in files:
        os.remove(f)

def branchwiseFullList(df):
    output_folder = "../output/full_branchwise"
    os.makedirs(output_folder, exist_ok=True)
    for branch in unique_branches:
        branch_df = df[df["Branch"] == branch][header]
        filePath = os.path.join(output_folder,f"{branch}.csv")
        branch_df.to_csv(filePath, index=False)

def branchMix(df):
    branches_data = {}
    folder = "../output/full_branchwise"
    for filename in os.listdir(folder):
        branch_name = os.path.splitext(filename)[0]
        df = pd.read_csv(os.path.join(folder, filename))
        branches_data[branch_name] = df.values.tolist()
    groupNum = 0
    output_folder = "../output/group_branch_wise_mix"
    os.makedirs(output_folder, exist_ok=True)
    while any(branches_data.values()):
        groupNum += 1
        result = []
        while len(result) < studPerGroup and any(branches_data.values()):
            for branch in branches_data:
                if branches_data[branch]:
                    result.append(branches_data[branch].pop(0))
                    if len(result) >= studPerGroup:
                        break
        file_path = os.path.join(output_folder, f"g{groupNum}.csv")
        group_df = pd.DataFrame(result, columns=header)
        group_df.to_csv(file_path, index=False)
        
        
def uniformMix(df):
    output_folder = "../output/group_uniform_mix"
    os.makedirs(output_folder, exist_ok=True)
    branches_data = {}
    folder = "../output/full_branchwise"
    for filename in os.listdir(folder):
        branch_name = os.path.splitext(filename)[0]
        df = pd.read_csv(os.path.join(folder, filename))
        branches_data[branch_name] = df.values.tolist()
    sorted_branches = sorted(branches_data.items(), key=lambda item: len(item[1]), reverse=True)
    all_students = []
    for branch_name, students in sorted_branches:
        all_students.extend(students)
    groupNum = 0
    while all_students:
        groupNum+=1
        result = []
        for _ in range(studPerGroup):
            if all_students:
                result.append(all_students.pop(0))
            else: break
        file_path = os.path.join(output_folder, f"g{groupNum}.csv")
        group_df = pd.DataFrame(result, columns=header)
        group_df.to_csv(file_path, index=False)
            
def makeOutputFile(folder_path,df):
    files = sorted(os.listdir(folder_path), key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))
    list_of_dicts = []
    total = []
    for filename in files:
        temp_dict = {}
        filepath = os.path.join(folder_path, filename)
        file_df = pd.read_csv(filepath)
        total.append(len(file_df))
        for branch in unique_branches:
            count = (file_df["Roll"].str.contains(branch)).sum()
            temp_dict[branch] = count
        list_of_dicts.append(temp_dict)
    df = pd.DataFrame(list_of_dicts)
    groups = [os.path.splitext(f)[0] for f in files if f.endswith(".csv")]
    if folder_path.endswith("full_branchwise"):
        df.insert(0, "MIX", groups)
    else:
        df.insert(0, "UNIFORM", groups)
    df["Total"] = total
    return df
    
        
                
st.title("Grouping Tool")
n = st.number_input("Enter the number of groups:", min_value=1, step=1)
inputFile = st.file_uploader("Input excel file", type=["xlsx"])

if inputFile is not None:
    input_path = os.path.join("../input",inputFile.name)
    with open(input_path,"wb") as f:
        f.write(inputFile.getbuffer())
    st.success(f"{inputFile.name} saved to input folder")
    df = pd.read_excel(inputFile, engine="openpyxl")
    
    st.subheader("Preview of input file")
    st.dataframe(df.head(10))
    df["Roll"] = df["Roll"].astype(str)
    df["Branch"] = df["Roll"].str[4:6]
    header = ["Roll","Name","Email"]
    unique_branches = df["Branch"].unique()
    studentsNum = len(df)
    studPerGroup = math.ceil(studentsNum / n)
    
if st.button("Start grouping"):
    if inputFile is not None and n>0:
        clear_folder("../output/full_branchwise")
        clear_folder("../output/group_branch_wise_mix")
        clear_folder("../output/group_uniform_mix")
        branchwiseFullList(df)
        branchMix(df)
        uniformMix(df)
        #output file code
        filepath = "../output.xlsx"
        if not os.path.exists(filepath):
            empty_df = pd.DataFrame()
            empty_df.to_excel(filepath, index=False)
        df1 = makeOutputFile("../output/group_branch_wise_mix",df)
        df2 = makeOutputFile("../output/group_uniform_mix",df)
        with pd.ExcelWriter("../output/output.xlsx") as writer:
            df1.to_excel(writer, sheet_name="Stats", index=False)
            start_row_df2 = len(df1) + 2 
            df2.to_excel(writer, sheet_name="Stats", startrow=start_row_df2, index=False)
            
        archive_name = "all_grouped_files"
        shutil.make_archive(archive_name, 'zip', '../output')
        zipName = f"{archive_name}.zip"
        with open(zipName, "rb") as f:
            st.download_button(
                label="Download All Results",
                data=f,
                file_name=zipName,
                mime="application/zip"
            )
    else:
        st.error("Please give number of groups and input file")
    

        





