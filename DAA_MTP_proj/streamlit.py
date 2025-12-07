# import streamlit as st
# import tempfile
# import os
# import shutil

# # Renaming imports using 'as' to change how they appear in this script
# from allocation import SeatingAllocator as CoreEngine
# from logs import setup_logging as init_logger

# def terminate_log_handlers(active_logger):
#     """
#     Detach and close file handlers to ensure Windows can delete the temp folder.
#     """
#     if not active_logger:
#         return

#     # Iterate over a copy of the handlers list to safely modify the original
#     current_handlers = list(active_logger.handlers)
#     for handler in current_handlers:
#         # Attempt to flush the stream
#         try:
#             handler.flush()
#         except Exception:
#             pass
        
#         # Attempt to close the file/stream
#         try:
#             handler.close()
#         except Exception:
#             pass
            
#         active_logger.removeHandler(handler)


# def execute_seating_process(source_file, margin_val, layout_mode):
#     # create a temporary workspace
#     with tempfile.TemporaryDirectory() as workspace:
#         # 1. Store the uploaded excel file locally
#         input_path = os.path.join(workspace, source_file.name)
#         with open(input_path, "wb") as dest:
#             dest.write(source_file.getbuffer())

#         # 2. Prepare the results directory
#         results_path = os.path.join(workspace, "output")
#         os.makedirs(results_path, exist_ok=True)

#         # 3. Initialize logging within the results folder
#         session_logger = init_logger(
#             logfile=os.path.join(results_path, "process_log.txt"),
#             errorfile=os.path.join(results_path, "error_report.txt"),
#         )

#         # 4. Initialize and run the core logic
#         # Note: 'CoreEngine' is the renamed SeatingAllocator class
#         engine = CoreEngine(
#             input_file=input_path,
#             buffer=margin_val,
#             density=layout_mode,
#             outdir=results_path,
#             logger=session_logger,
#         )

#         # These method names must match functions in allocation.py
#         engine.load_inputs()
#         engine.allocate_all_days()
#         engine.write_outputs()

#         # 5. Handle Image/PDF generation
#         img_folder = "photos" 
#         placeholder_img = os.path.join(img_folder, "no_image_available.jpg")
        
#         # This calls the PDF generator in allocation.py
#         engine.generate_attendance_pdfs(img_folder, placeholder_img)

#         # 6. Compress the output for download
#         archive_name = os.path.join(workspace, "final_attendance_pack")
#         shutil.make_archive(archive_name, "zip", results_path)
#         final_zip_path = archive_name + ".zip"

#         # 7. Read the zip file into memory
#         with open(final_zip_path, "rb") as zf:
#             archive_bytes = zf.read()

#         # 8. Clean up logger so Windows allows folder deletion
#         terminate_log_handlers(session_logger)

#         return archive_bytes


# # --- Streamlit UI Setup ---

# st.title("Exam Allocation & Attendance Tool")

# # Renamed UI variables
# raw_excel = st.file_uploader("Select your Excel Data Sheet", type=["xlsx"])
# seat_margin = st.number_input("Empty seats between students (Buffer)", min_value=0, max_value=50, value=0)
# mode_select = st.selectbox("Arrangement Mode", ["Dense", "Sparse"])

# # Logic Trigger
# if st.button("Start Processing") and raw_excel:
#     with st.spinner("Generating arrangement, please wait..."):
#         try:
#             # Call the renamed function
#             final_data = execute_seating_process(raw_excel, seat_margin, mode_select)
            
#             st.success("Processing Complete!")
#             st.download_button(
#                 label="Download Results (ZIP)",
#                 data=final_data,
#                 file_name="exam_seating_results.zip",
#                 mime="application/zip",
#             )
#         except Exception as err:
#             st.error(f"An unexpected error occurred: {err}")




import streamlit as st
import tempfile
import os
import shutil

# IMPORT MATCH: Importing the NEW class name defined in allocation.py
from allocation import ExamScheduler as CoreEngine
# IMPORT MATCH: Importing the NEW function name defined in logs.py
from logs import init_logger

def terminate_log_handlers(active_logger):
    """
    Detach and close file handlers to ensure Windows can delete the temp folder.
    """
    if not active_logger:
        return

    # Iterate over a copy of the handlers list to safely modify the original
    current_handlers = list(active_logger.handlers)
    for handler in current_handlers:
        try:
            handler.flush()
        except Exception:
            pass
        
        try:
            handler.close()
        except Exception:
            pass
            
        active_logger.removeHandler(handler)


def inject_custom_ui():
    """
    Applies custom CSS for a modern, professional look.
    """
    st.markdown("""
    <style>
        /* Modern Button Style */
        div.stButton > button {
            background: linear-gradient(90deg, #4b6cb7 0%, #182848 100%);
            color: white;
            border: none;
            padding: 0.6rem 1.2rem;
            border-radius: 8px;
            font-size: 1rem;
            transition: all 0.3s ease;
            width: 100%;
        }
        div.stButton > button:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            transform: translateY(-2px);
            color: #ffffff;
        }
        
        /* Sidebar styling */
        [data-testid="stSidebar"] {
            background-color: #f8f9fa;
        }
        
        /* Headers */
        h1 {
            color: #182848;
            font-family: 'Segoe UI', sans-serif;
            font-weight: 700;
        }
        h2, h3 {
            color: #4b6cb7;
        }
        
        /* Success Message Box */
        .stSuccess {
            background-color: #d4edda;
            color: #155724;
            border-color: #c3e6cb;
        }
    </style>
    """, unsafe_allow_html=True)


def execute_seating_process(source_file, margin_val, layout_mode):
    # create a temporary workspace
    with tempfile.TemporaryDirectory() as workspace:
        # 1. Store the uploaded excel file locally
        input_path = os.path.join(workspace, source_file.name)
        with open(input_path, "wb") as dest:
            dest.write(source_file.getbuffer())

        # 2. Prepare the results directory
        results_path = os.path.join(workspace, "output")
        os.makedirs(results_path, exist_ok=True)

        # 3. Initialize logging within the results folder
        session_logger = init_logger(
            log_path=os.path.join(results_path, "process_log.txt"),
            error_path=os.path.join(results_path, "error_report.txt"),
        )

        # 4. Initialize and run the core logic
        # Note: 'CoreEngine' is the renamed ExamScheduler class from allocation.py
        # MATCHING PARAMETERS: These must match the __init__ in allocation.py
        engine = CoreEngine(
            source_path=input_path,
            gap_size=margin_val,
            layout_mode=layout_mode,
            result_dir=results_path,
            log_handler=session_logger,
        )

        # MATCHING METHODS: These must match the RENAMED functions in allocation.py
        # (Using the new names I provided in the previous steps)
        engine.load_and_parse_data()
        engine.process_all_slots()
        engine.generate_excel_reports()

        # 5. Handle Image/PDF generation
        img_folder = "photos" 
        placeholder_img = os.path.join(img_folder, "no_image_available.jpg")
        
        # MATCHING METHOD: This must match the RENAMED function in allocation.py
        engine.create_attendance_files(img_folder, placeholder_img)

        # 6. Compress the output for download
        archive_name = os.path.join(workspace, "final_attendance_pack")
        shutil.make_archive(archive_name, "zip", results_path)
        final_zip_path = archive_name + ".zip"

        # 7. Read the zip file into memory
        with open(final_zip_path, "rb") as zf:
            archive_bytes = zf.read()

        # 8. Clean up logger so Windows allows folder deletion
        terminate_log_handlers(session_logger)

        return archive_bytes


# --- Streamlit UI Setup ---

# Page Config
st.set_page_config(
    page_title="Exam Manager",
    page_icon="🏛️",
    layout="wide"
)

# Inject Custom CSS
inject_custom_ui()

# --- Sidebar ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2232/2232688.png", width=80) # Generic exam icon
    st.title("Settings")
    st.markdown("Configure your seating logic here.")
    
    st.divider()
    
    seat_margin = st.slider(
        "🛑 Buffer Seats (Gap)", 
        min_value=0, 
        max_value=20, 
        value=0,
        help="Number of empty seats to leave per room."
    )
    
    mode_select = st.radio(
        "🪑 Seating Density", 
        ["Dense", "Sparse"],
        help="Dense fills rooms; Sparse skips every other seat."
    )
    
    st.divider()
    st.info("System Ready v2.0")

# --- Main Content ---

st.title("🏛️ Exam Seating & Attendance Manager")
st.markdown("Upload your master schedule to generate room allocations and attendance sheets automatically.")

# Instructions Expander
with st.expander("ℹ️ How to use this tool"):
    st.markdown("""
    1. **Prepare Data:** Ensure your Excel file has `in_timetable`, `in_room_capacity`, etc.
    2. **Upload:** Drag and drop the file below.
    3. **Configure:** Adjust buffer and density in the sidebar.
    4. **Generate:** Click the button to process and download the ZIP.
    """)

# Main Layout
col1, col2 = st.columns([2, 1])

with col1:
    raw_excel = st.file_uploader("📂 Upload Master Excel Sheet", type=["xlsx"])

if raw_excel:
    with col2:
        st.write("##") # Spacer
        st.write("##") # Spacer
        # Action Button
        start_btn = st.button("⚡ Generate Arrangement")

    if start_btn:
        # Progress Bar and Status
        with st.status("Processing data...", expanded=True) as status:
            try:
                st.write("Reading Excel sheets...")
                # Call the processing function
                final_data = execute_seating_process(raw_excel, seat_margin, mode_select)
                
                st.write("Allocating students...")
                st.write("Generating PDF files...")
                st.write("Zipping output...")
                status.update(label="Process Complete!", state="complete", expanded=False)
                
                st.divider()
                st.success("✅ Allocation successful!")
                
                # Download Button
                st.download_button(
                    label="📥 Download Final Package (.zip)",
                    data=final_data,
                    file_name="exam_seating_output.zip",
                    mime="application/zip",
                )
                
            except Exception as err:
                status.update(label="Error Occurred", state="error")
                st.error(f"❌ System Error: {err}")
else:
    # Empty state helper
    st.info("Please upload a file to begin.")