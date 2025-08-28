# README.txt  

Requirements
Make sure you have Python 3.9+ installed.  

Install dependencies from `requirements.txt`:  
```
pip install -r requirements.txt
```


How to Execute the Code

1. **Prepare the input file**  
   - Place your Excel input file (e.g., `students.xlsx`) in the **input/** folder.  
   - The Excel file must contain these columns:  
     - Roll (student roll number)  
     - Name (student name)  
     - Email (student email)  

2. **Run the Streamlit app**  
   Open a terminal inside the project root and run:  
   ```
   streamlit run How_to_execute/tut01.py
   ```

3. **Use the web app**  
   - A browser window will open automatically (default: http://localhost:8501).  
   - In the app:  
     - Enter the number of groups.  
     - Upload your Excel input file.  
     - Click **Start grouping**.  

4. **Output**  
   - Grouped files are saved in the **output/** folder, under:  
     - full_branchwise/  
     - group_branch_wise_mix/  
     - group_uniform_mix/  
   - A combined summary Excel file is saved as output/output.xlsx.  
   - You can also download all results as a ZIP file directly from the app.  

---


