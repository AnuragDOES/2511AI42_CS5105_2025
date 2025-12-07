# """
# allocation.py — Core logic for exam seating arrangement generation (Refactored).
# """

# import os
# import math
# import pandas as pd
# from collections import defaultdict
# # Make sure your utility file is actually named 'read_write_utils.py' or update this import
# from read_write_utils import read_excel_file 
# from attendance import build_attendance_pdf

# class ExamScheduler:
#     def __init__(self, source_path, gap_size=0, layout_mode='Dense', result_dir='output', log_handler=None):
#         self.src_excel = source_path
#         self.gap_size = int(gap_size)
#         self.layout_mode = layout_mode  # 'Dense' or 'Sparse'
#         self.result_dir = result_dir
#         self.log = log_handler

#         # Data Containers
#         self.raw_sheets = {}
#         self.exam_schedule = None  # List of days/slots
#         self.student_courses = None  # DataFrame of registrations
#         self.student_names = {}  # Roll -> Name
#         self.course_enrollment = defaultdict(list)  # Subject -> List of Rolls
#         self.venues = []  # List of room dictionaries
#         self.final_assignments = defaultdict(list)  # Key -> List of allocated students

#         os.makedirs(self.result_dir, exist_ok=True)

#     def load_and_parse_data(self):
#         """
#         Ingest data from the excel source:
#            - in_timetable
#            - in_course_roll_mapping
#            - in_roll_name_mapping
#            - in_room_capacity
#         """
#         try:
#             self.log.info("Reading source file: %s", self.src_excel)
#             self.raw_sheets = read_excel_file(self.src_excel, logger=self.log)

#             # 1. Process Timetable
#             if 'in_timetable' not in self.raw_sheets:
#                 raise ValueError("Sheet missing: in_timetable")

#             df_schedule = self.raw_sheets['in_timetable']
#             needed_cols = ['Date', 'Day', 'Morning', 'Evening']
#             for c in needed_cols:
#                 if c not in df_schedule.columns:
#                     raise ValueError(f"in_timetable needs column: {c}")

#             self.exam_schedule = []
#             for _, row in df_schedule.iterrows():
#                 dt = str(row['Date']).strip()
#                 dy = str(row['Day']).strip()

#                 def _split_subjects(val):
#                     if pd.isna(val): 
#                         return ['NO EXAM']
#                     s_val = str(val).strip()
#                     if not s_val or s_val.upper() == 'NO EXAM':
#                         return ['NO EXAM']
#                     return [x.strip() for x in s_val.split(';') if x.strip()]

#                 s1_courses = _split_subjects(row['Morning'])
#                 s2_courses = _split_subjects(row['Evening'])

#                 self.exam_schedule.append({
#                     'Date': dt,
#                     'Day': dy,
#                     'Morning': s1_courses,
#                     'Evening': s2_courses
#                 })
#             self.log.info("Schedule loaded: %d days found.", len(self.exam_schedule))

#             # 2. Process Names (Optional)
#             if 'in_roll_name_mapping' in self.raw_sheets:
#                 df_n = self.raw_sheets['in_roll_name_mapping']
#                 # Case-insensitive column search
#                 low_cols = {x.lower(): x for x in df_n.columns}
                
#                 if 'roll' in low_cols and 'name' in low_cols:
#                     col_r = low_cols['roll']
#                     col_n = low_cols['name']
#                     for _, r in df_n.iterrows():
#                         rl = str(r[col_r]).strip()
#                         nm = str(r[col_n]).strip() or 'Unknown'
#                         if rl:
#                             self.student_names[rl] = nm
#                 else:
#                     self.log.warning("Name mapping sheet found but columns (Roll, Name) missing.")
#             else:
#                 self.log.warning("No name mapping sheet found.")

#             # 3. Process Registrations
#             if 'in_course_roll_mapping' not in self.raw_sheets:
#                 raise ValueError("Sheet missing: in_course_roll_mapping")
            
#             df_reg = self.raw_sheets['in_course_roll_mapping']
#             # validation
#             l_cols = {c.lower(): c for c in df_reg.columns}
#             if 'rollno' in l_cols and 'course_code' in l_cols:
#                 c_roll = l_cols['rollno']
#                 c_code = l_cols['course_code']
#             else:
#                 raise ValueError("Mapping sheet needs 'rollno' and 'course_code'")

#             self.student_courses = df_reg
#             rec_count = 0
#             for _, r in df_reg.iterrows():
#                 rl = str(r[c_roll]).strip()
#                 sb = str(r[c_code]).strip()
#                 if rl and sb:
#                     self.course_enrollment[sb].append(rl)
#                     rec_count += 1
#             self.log.info("Registrations loaded: %d entries.", rec_count)

#             # 4. Process Rooms
#             if 'in_room_capacity' not in self.raw_sheets:
#                 raise ValueError("Sheet missing: in_room_capacity")
            
#             df_rooms = self.raw_sheets['in_room_capacity']
#             col_map = {c.strip().lower(): c for c in df_rooms.columns}
            
#             # Check mandatory columns
#             if not all(k in col_map for k in ['room no.', 'exam capacity', 'block']):
#                  raise ValueError("Room sheet needs: Room No., Exam Capacity, Block")

#             self.venues = []
#             for _, r in df_rooms.iterrows():
#                 r_code = str(r[col_map['room no.']]).strip()
#                 try:
#                     raw_cap = int(float(r[col_map['exam capacity']]))
#                 except:
#                     raw_cap = 0
                
#                 blk = str(r[col_map['block']]).strip()
#                 real_cap = self._calc_seat_limit(raw_cap)
                
#                 self.venues.append({
#                     'building': blk,
#                     'room_code': r_code,
#                     'capacity': raw_cap,
#                     'capacity_effective': real_cap
#                 })
#             self.log.info("Venues loaded: %d rooms.", len(self.venues))

#         except Exception as ex:
#             self.log.exception("Data import failed: %s", ex)
#             raise

#     def _calc_seat_limit(self, total_seats):
#         """Calculate usable seats based on buffer and density settings."""
#         try:
#             val = max(0, int(total_seats) - self.gap_size)
#         except:
#             val = 0
        
#         if str(self.layout_mode).strip().lower() == 'sparse':
#             return val // 2
#         return val

#     def scan_for_conflicts(self):
#         """Ensure no student has two exams in the same slot."""
#         try:
#             if self.student_courses is None:
#                 raise ValueError("No student data available for clash check.")

#             df = self.student_courses
#             lc = {c.lower(): c for c in df.columns}
#             col_r = lc['rollno']
#             col_c = lc['course_code']

#             has_issue = False

#             for day_entry in self.exam_schedule:
#                 dt = day_entry['Date']
#                 for slot, sub_list in [('Morning', day_entry['Morning']), ('Evening', day_entry['Evening'])]:
#                     if sub_list == ['NO EXAM']:
#                         continue

#                     # map subject -> set of students
#                     sub_map = {}
#                     for sb in sub_list:
#                         clean_sb = str(sb).strip()
#                         # filter dataframe
#                         mask = df[col_c].astype(str).str.strip() == clean_sb
#                         stds = set(str(x).strip() for x in df.loc[mask, col_r].dropna())
#                         sub_map[clean_sb] = stds

#                     # Check intersections
#                     all_subs = list(sub_map.keys())
#                     for i in range(len(all_subs)):
#                         for j in range(i + 1, len(all_subs)):
#                             s1, s2 = all_subs[i], all_subs[j]
#                             common = sub_map[s1] & sub_map[s2]
#                             if common:
#                                 has_issue = True
#                                 for std in common:
#                                     self.log.error(f"CLASH detected: {dt} ({slot}) | {s1} & {s2} | Student: {std}")

#             if has_issue:
#                 print("⚠️ Clashes found! Check log.")
#             else:
#                 self.log.info("Validation passed: No clashes.")

#         except Exception as e:
#             self.log.exception("Conflict scan failed: %s", e)
#             raise

#     def assign_students_to_rooms(self, subject_code, student_list, available_rooms):
#         """
#         Greedy allocation: Fill largest rooms first.
#         Returns: (list_of_assignments, remaining_students)
#         """
#         try:
#             output = []
#             queue = list(student_list)
            
#             # Sort rooms by space left (High to Low)
#             prioritized_rooms = sorted(available_rooms, key=lambda x: x.get('capacity_effective', 0), reverse=True)

#             for rm in prioritized_rooms:
#                 if not queue:
#                     break
                
#                 limit = int(rm.get('capacity_effective', 0))
#                 if limit <= 0:
#                     continue

#                 count_to_take = min(len(queue), limit)
#                 batch = queue[:count_to_take]
#                 queue = queue[count_to_take:]

#                 output.append({
#                     'building': rm.get('building'),
#                     'room': rm.get('room_code'),
#                     'rolls': batch
#                 })
            
#             return output, queue
#         except Exception as e:
#             self.log.exception("Allocation error for %s: %s", subject_code, e)
#             raise

#     def process_all_slots(self):
#         """Run the main scheduling logic for every day/slot."""
#         try:
#             self.scan_for_conflicts()

#             for day_data in self.exam_schedule:
#                 date_val = day_data['Date']
#                 day_str = day_data['Day']
                
#                 # Format folder name
#                 d_clean = str(date_val).split()[0].replace("-", "_").replace("/", "_")
#                 base_path = os.path.join(self.result_dir, d_clean)
                
#                 paths = {
#                     'Morning': os.path.join(base_path, 'Morning'),
#                     'Evening': os.path.join(base_path, 'Evening')
#                 }
#                 for p in paths.values():
#                     os.makedirs(p, exist_ok=True)

#                 # Process Morning and Evening
#                 for slot, subjects in [('Morning', day_data['Morning']), ('Evening', day_data['Evening'])]:
#                     curr_dir = paths[slot]

#                     if subjects == ['NO EXAM']:
#                         with open(os.path.join(curr_dir, 'NO_EXAM.txt'), 'w', encoding='utf-8') as f:
#                             f.write('NO EXAM')
#                         continue

#                     # Reset room capacities for this slot
#                     slot_venues = [dict(v) for v in self.venues]

#                     # Sort subjects by student count (Largest first)
#                     weighted_subs = [(s, len(self.course_enrollment.get(s, []))) for s in subjects]
#                     weighted_subs.sort(key=lambda x: x[1], reverse=True)

#                     for sub_code, _ in weighted_subs:
#                         sub_code = str(sub_code).strip()
#                         students = self.course_enrollment.get(sub_code, [])

#                         if not students:
#                             self.log.warning("Empty student list for %s", sub_code)
#                             # Create empty placeholder file
#                             pd.DataFrame(columns=['Room', 'Rolls', 'Count']).to_excel(
#                                 os.path.join(curr_dir, f"{sub_code}.xlsx"), index=False
#                             )
#                             continue

#                         # Perform allocation
#                         assignments, unassigned = self.assign_students_to_rooms(sub_code, students, slot_venues)

#                         if unassigned:
#                             msg = f"Overflow: {len(unassigned)} students could not fit for {sub_code} on {date_val}"
#                             self.log.error(msg)
#                             raise RuntimeError(msg)

#                         # Commit updates
#                         for item in assignments:
#                             # 1. Reduce room capacity
#                             for v in slot_venues:
#                                 if v['room_code'] == item['room'] and v['building'] == item['building']:
#                                     v['capacity_effective'] = max(0, v.get('capacity_effective', 0) - len(item['rolls']))
#                                     break
                            
#                             # 2. Save to global registry
#                             key = f"{date_val}_{slot}"
#                             self.final_assignments[key].append({
#                                 'date': date_val,
#                                 'day': day_str,
#                                 'slot': slot,
#                                 'subject': sub_code,
#                                 'building': item['building'],
#                                 'room': item['room'],
#                                 'rolls': item['rolls']
#                             })

#                         # Write per-subject file
#                         rows = []
#                         for item in assignments:
#                             rows.append({
#                                 'Room': item['room'],
#                                 'Rolls (semicolon separated)': ';'.join(item['rolls']),
#                                 'Count': len(item['rolls'])
#                             })
#                         pd.DataFrame(rows).to_excel(os.path.join(curr_dir, f"{sub_code}.xlsx"), index=False)

#                     self.log.info("Completed slot: %s %s", date_val, slot)

#         except Exception as e:
#             self.log.exception("Scheduling failed: %s", e)
#             raise

#     def generate_excel_reports(self):
#         """Create the master summary and seats-left report."""
#         try:
#             # 1. Master Allocation Report
#             master_data = []
#             for _, alloc_list in self.final_assignments.items():
#                 for entry in alloc_list:
#                     master_data.append({
#                         "Date": entry["date"],
#                         "Day": entry.get("day", ""),
#                         "course_code": entry["subject"],
#                         "Room": entry["room"],
#                         "Allocated_students_count": len(entry["rolls"]),
#                         "Roll_list (semicolon separated)": ";".join(entry["rolls"]),
#                     })

#             out_1 = os.path.join(self.result_dir, "op_overall_seating_arrangement.xlsx")
#             pd.DataFrame(master_data).to_excel(out_1, index=False)

#             # 2. Vacancy Report (Seats Left)
#             grouped_by_slot = defaultdict(list)
#             for _, alloc_list in self.final_assignments.items():
#                 for entry in alloc_list:
#                     k = (str(entry["date"]), str(entry["slot"]))
#                     grouped_by_slot[k].append(entry)

#             out_2 = os.path.join(self.result_dir, "op_seats_left.xlsx")
            
#             with pd.ExcelWriter(out_2, engine="xlsxwriter") as writer:
#                 for (d, s), items in grouped_by_slot.items():
#                     # Tally usage per room
#                     usage_map = {v["room_code"]: 0 for v in self.venues}
#                     for i in items:
#                         usage_map[i["room"]] = usage_map.get(i["room"], 0) + len(i["rolls"])

#                     report_rows = []
#                     for v in self.venues:
#                         used = usage_map.get(v["room_code"], 0)
#                         free = max(0, v["capacity"] - used)
#                         report_rows.append({
#                             "Room No.": v["room_code"],
#                             "Exam Capacity": v["capacity"],
#                             "Block": v["building"],
#                             "Alloted": used,
#                             "Vacant (B-C)": free,
#                         })

#                     # Safe sheet name
#                     s_name = f"{str(d).split()[0].replace('-', '_')}_{s}"
#                     for bad in '[]:*?/\\':
#                         s_name = s_name.replace(bad, "_")
                    
#                     pd.DataFrame(report_rows).to_excel(writer, sheet_name=s_name[:31], index=False)

#             self.log.info("Reports saved to %s and %s", out_1, out_2)

#         except Exception as e:
#             self.log.exception("Report generation failed: %s", e)
#             raise

#     def create_attendance_files(self, img_source, default_icon, pdf_root=None):
#         """Generate PDF attendance sheets for every room/subject combination."""
#         if pdf_root is None:
#             pdf_root = os.path.join(self.result_dir, "attendance")
#         os.makedirs(pdf_root, exist_ok=True)
        
#         self.log.info("Starting PDF generation in: %s", pdf_root)

#         # Re-organize data for printing
#         print_queue = {}
#         for _, alloc_list in self.final_assignments.items():
#             for item in alloc_list:
#                 key = (str(item["date"]), str(item["slot"]), str(item["room"]), str(item["subject"]))
#                 print_queue.setdefault(key, []).extend(item["rolls"])

#         for (dt, sl, rm, sb), roll_list in print_queue.items():
#             unique_rolls = list(dict.fromkeys(roll_list))
            
#             # Sanitize filename
#             d_simple = str(dt).split()[0].replace("-", "_").replace("/", "_")
#             f_name = f"{d_simple}_{sl}_{rm}_{sb}.pdf"
#             for bad in '<>:"/\\|?* ':
#                 f_name = f_name.replace(bad, "_")
            
#             full_path = os.path.join(pdf_root, f_name)
            
#             try:
#                 build_attendance_pdf(
#                     out_path=full_path,
#                     date_str=str(dt).split()[0],
#                     shift=sl,
#                     room_no=rm,
#                     subject_code=sb,
#                     subject_name=sb,
#                     roll_list=unique_rolls,
#                     roll_to_name=self.student_names,
#                     photos_dir=img_source,
#                     no_image_icon=default_icon,
#                     logger=self.log,
#                 )
#             except Exception:
#                 self.log.error(f"Failed to create PDF for {sb} in {rm}")
#                 continue

#         self.log.info("PDF generation complete.")



"""
allocation.py — Core logic for exam seating arrangement generation (Refactored).
"""

import os
import math
import pandas as pd
from collections import defaultdict
# IMPORT UPDATE: Ensure your file is named 'read_write_utils.py' (formerly io.py)
from read_write_utils import read_excel_file 
from attendance import build_attendance_pdf

class ExamScheduler:
    def __init__(self, source_path, gap_size=0, layout_mode='Dense', result_dir='output', log_handler=None):
        self.src_excel = source_path
        self.gap_size = int(gap_size)
        self.layout_mode = layout_mode  # 'Dense' or 'Sparse'
        self.result_dir = result_dir
        self.log = log_handler

        # Data Containers
        self.raw_sheets = {}
        self.exam_schedule = None  # List of days/slots
        self.student_courses = None  # DataFrame of registrations
        self.student_names = {}  # Roll -> Name
        self.course_enrollment = defaultdict(list)  # Subject -> List of Rolls
        self.venues = []  # List of room dictionaries
        self.final_assignments = defaultdict(list)  # Key -> List of allocated students

        os.makedirs(self.result_dir, exist_ok=True)

    def load_and_parse_data(self):
        """
        Ingest data from the excel source:
           - in_timetable
           - in_course_roll_mapping
           - in_roll_name_mapping
           - in_room_capacity
        """
        try:
            self.log.info("Reading source file: %s", self.src_excel)
            self.raw_sheets = read_excel_file(self.src_excel, logger=self.log)

            # 1. Process Timetable
            if 'in_timetable' not in self.raw_sheets:
                raise ValueError("Sheet missing: in_timetable")

            df_schedule = self.raw_sheets['in_timetable']
            needed_cols = ['Date', 'Day', 'Morning', 'Evening']
            for c in needed_cols:
                if c not in df_schedule.columns:
                    raise ValueError(f"in_timetable needs column: {c}")

            self.exam_schedule = []
            for _, row in df_schedule.iterrows():
                dt = str(row['Date']).strip()
                dy = str(row['Day']).strip()

                def _split_subjects(val):
                    if pd.isna(val): 
                        return ['NO EXAM']
                    s_val = str(val).strip()
                    if not s_val or s_val.upper() == 'NO EXAM':
                        return ['NO EXAM']
                    return [x.strip() for x in s_val.split(';') if x.strip()]

                s1_courses = _split_subjects(row['Morning'])
                s2_courses = _split_subjects(row['Evening'])

                self.exam_schedule.append({
                    'Date': dt,
                    'Day': dy,
                    'Morning': s1_courses,
                    'Evening': s2_courses
                })
            self.log.info("Schedule loaded: %d days found.", len(self.exam_schedule))

            # 2. Process Names (Optional)
            if 'in_roll_name_mapping' in self.raw_sheets:
                df_n = self.raw_sheets['in_roll_name_mapping']
                # Case-insensitive column search
                low_cols = {x.lower(): x for x in df_n.columns}
                
                if 'roll' in low_cols and 'name' in low_cols:
                    col_r = low_cols['roll']
                    col_n = low_cols['name']
                    for _, r in df_n.iterrows():
                        rl = str(r[col_r]).strip()
                        nm = str(r[col_n]).strip() or 'Unknown'
                        if rl:
                            self.student_names[rl] = nm
                else:
                    self.log.warning("Name mapping sheet found but columns (Roll, Name) missing.")
            else:
                self.log.warning("No name mapping sheet found.")

            # 3. Process Registrations
            if 'in_course_roll_mapping' not in self.raw_sheets:
                raise ValueError("Sheet missing: in_course_roll_mapping")
            
            df_reg = self.raw_sheets['in_course_roll_mapping']
            # validation
            l_cols = {c.lower(): c for c in df_reg.columns}
            if 'rollno' in l_cols and 'course_code' in l_cols:
                c_roll = l_cols['rollno']
                c_code = l_cols['course_code']
            else:
                raise ValueError("Mapping sheet needs 'rollno' and 'course_code'")

            self.student_courses = df_reg
            rec_count = 0
            for _, r in df_reg.iterrows():
                rl = str(r[c_roll]).strip()
                sb = str(r[c_code]).strip()
                if rl and sb:
                    self.course_enrollment[sb].append(rl)
                    rec_count += 1
            self.log.info("Registrations loaded: %d entries.", rec_count)

            # 4. Process Rooms
            if 'in_room_capacity' not in self.raw_sheets:
                raise ValueError("Sheet missing: in_room_capacity")
            
            df_rooms = self.raw_sheets['in_room_capacity']
            col_map = {c.strip().lower(): c for c in df_rooms.columns}
            
            # Check mandatory columns
            if not all(k in col_map for k in ['room no.', 'exam capacity', 'block']):
                 raise ValueError("Room sheet needs: Room No., Exam Capacity, Block")

            self.venues = []
            for _, r in df_rooms.iterrows():
                r_code = str(r[col_map['room no.']]).strip()
                try:
                    raw_cap = int(float(r[col_map['exam capacity']]))
                except:
                    raw_cap = 0
                
                blk = str(r[col_map['block']]).strip()
                real_cap = self._calc_seat_limit(raw_cap)
                
                self.venues.append({
                    'building': blk,
                    'room_code': r_code,
                    'capacity': raw_cap,
                    'capacity_effective': real_cap
                })
            self.log.info("Venues loaded: %d rooms.", len(self.venues))

        except Exception as ex:
            self.log.exception("Data import failed: %s", ex)
            raise

    def _calc_seat_limit(self, total_seats):
        """Calculate usable seats based on buffer and density settings."""
        try:
            val = max(0, int(total_seats) - self.gap_size)
        except:
            val = 0
        
        if str(self.layout_mode).strip().lower() == 'sparse':
            return val // 2
        return val

    def scan_for_conflicts(self):
        """Ensure no student has two exams in the same slot."""
        try:
            if self.student_courses is None:
                raise ValueError("No student data available for clash check.")

            df = self.student_courses
            lc = {c.lower(): c for c in df.columns}
            col_r = lc['rollno']
            col_c = lc['course_code']

            has_issue = False

            for day_entry in self.exam_schedule:
                dt = day_entry['Date']
                for slot, sub_list in [('Morning', day_entry['Morning']), ('Evening', day_entry['Evening'])]:
                    if sub_list == ['NO EXAM']:
                        continue

                    # map subject -> set of students
                    sub_map = {}
                    for sb in sub_list:
                        clean_sb = str(sb).strip()
                        # filter dataframe
                        mask = df[col_c].astype(str).str.strip() == clean_sb
                        stds = set(str(x).strip() for x in df.loc[mask, col_r].dropna())
                        sub_map[clean_sb] = stds

                    # Check intersections
                    all_subs = list(sub_map.keys())
                    for i in range(len(all_subs)):
                        for j in range(i + 1, len(all_subs)):
                            s1, s2 = all_subs[i], all_subs[j]
                            common = sub_map[s1] & sub_map[s2]
                            if common:
                                has_issue = True
                                for std in common:
                                    self.log.error(f"CLASH detected: {dt} ({slot}) | {s1} & {s2} | Student: {std}")

            if has_issue:
                print("⚠️ Clashes found! Check log.")
            else:
                self.log.info("Validation passed: No clashes.")

        except Exception as e:
            self.log.exception("Conflict scan failed: %s", e)
            raise

    def assign_students_to_rooms(self, subject_code, student_list, available_rooms):
        """
        Greedy allocation: Fill largest rooms first.
        Returns: (list_of_assignments, remaining_students)
        """
        try:
            output = []
            queue = list(student_list)
            
            # Sort rooms by space left (High to Low)
            prioritized_rooms = sorted(available_rooms, key=lambda x: x.get('capacity_effective', 0), reverse=True)

            for rm in prioritized_rooms:
                if not queue:
                    break
                
                limit = int(rm.get('capacity_effective', 0))
                if limit <= 0:
                    continue

                count_to_take = min(len(queue), limit)
                batch = queue[:count_to_take]
                queue = queue[count_to_take:]

                output.append({
                    'building': rm.get('building'),
                    'room': rm.get('room_code'),
                    'rolls': batch
                })
            
            return output, queue
        except Exception as e:
            self.log.exception("Allocation error for %s: %s", subject_code, e)
            raise

    def process_all_slots(self):
        """Run the main scheduling logic for every day/slot."""
        try:
            self.scan_for_conflicts()

            for day_data in self.exam_schedule:
                date_val = day_data['Date']
                day_str = day_data['Day']
                
                # Format folder name
                d_clean = str(date_val).split()[0].replace("-", "_").replace("/", "_")
                base_path = os.path.join(self.result_dir, d_clean)
                
                paths = {
                    'Morning': os.path.join(base_path, 'Morning'),
                    'Evening': os.path.join(base_path, 'Evening')
                }
                for p in paths.values():
                    os.makedirs(p, exist_ok=True)

                # Process Morning and Evening
                for slot, subjects in [('Morning', day_data['Morning']), ('Evening', day_data['Evening'])]:
                    curr_dir = paths[slot]

                    if subjects == ['NO EXAM']:
                        with open(os.path.join(curr_dir, 'NO_EXAM.txt'), 'w', encoding='utf-8') as f:
                            f.write('NO EXAM')
                        continue

                    # Reset room capacities for this slot
                    slot_venues = [dict(v) for v in self.venues]

                    # Sort subjects by student count (Largest first)
                    weighted_subs = [(s, len(self.course_enrollment.get(s, []))) for s in subjects]
                    weighted_subs.sort(key=lambda x: x[1], reverse=True)

                    for sub_code, _ in weighted_subs:
                        sub_code = str(sub_code).strip()
                        students = self.course_enrollment.get(sub_code, [])

                        if not students:
                            self.log.warning("Empty student list for %s", sub_code)
                            # Create empty placeholder file
                            pd.DataFrame(columns=['Room', 'Rolls', 'Count']).to_excel(
                                os.path.join(curr_dir, f"{sub_code}.xlsx"), index=False
                            )
                            continue

                        # Perform allocation
                        assignments, unassigned = self.assign_students_to_rooms(sub_code, students, slot_venues)

                        if unassigned:
                            msg = f"Overflow: {len(unassigned)} students could not fit for {sub_code} on {date_val}"
                            self.log.error(msg)
                            raise RuntimeError(msg)

                        # Commit updates
                        for item in assignments:
                            # 1. Reduce room capacity
                            for v in slot_venues:
                                if v['room_code'] == item['room'] and v['building'] == item['building']:
                                    v['capacity_effective'] = max(0, v.get('capacity_effective', 0) - len(item['rolls']))
                                    break
                            
                            # 2. Save to global registry
                            key = f"{date_val}_{slot}"
                            self.final_assignments[key].append({
                                'date': date_val,
                                'day': day_str,
                                'slot': slot,
                                'subject': sub_code,
                                'building': item['building'],
                                'room': item['room'],
                                'rolls': item['rolls']
                            })

                        # Write per-subject file
                        rows = []
                        for item in assignments:
                            rows.append({
                                'Room': item['room'],
                                'Rolls (semicolon separated)': ';'.join(item['rolls']),
                                'Count': len(item['rolls'])
                            })
                        pd.DataFrame(rows).to_excel(os.path.join(curr_dir, f"{sub_code}.xlsx"), index=False)

                    self.log.info("Completed slot: %s %s", date_val, slot)

        except Exception as e:
            self.log.exception("Scheduling failed: %s", e)
            raise

    def generate_excel_reports(self):
        """Create the master summary and seats-left report."""
        try:
            # 1. Master Allocation Report
            master_data = []
            for _, alloc_list in self.final_assignments.items():
                for entry in alloc_list:
                    master_data.append({
                        "Date": entry["date"],
                        "Day": entry.get("day", ""),
                        "course_code": entry["subject"],
                        "Room": entry["room"],
                        "Allocated_students_count": len(entry["rolls"]),
                        "Roll_list (semicolon separated)": ";".join(entry["rolls"]),
                    })

            out_1 = os.path.join(self.result_dir, "op_overall_seating_arrangement.xlsx")
            pd.DataFrame(master_data).to_excel(out_1, index=False)

            # 2. Vacancy Report (Seats Left)
            grouped_by_slot = defaultdict(list)
            for _, alloc_list in self.final_assignments.items():
                for entry in alloc_list:
                    k = (str(entry["date"]), str(entry["slot"]))
                    grouped_by_slot[k].append(entry)

            out_2 = os.path.join(self.result_dir, "op_seats_left.xlsx")
            
            with pd.ExcelWriter(out_2, engine="xlsxwriter") as writer:
                for (d, s), items in grouped_by_slot.items():
                    # Tally usage per room
                    usage_map = {v["room_code"]: 0 for v in self.venues}
                    for i in items:
                        usage_map[i["room"]] = usage_map.get(i["room"], 0) + len(i["rolls"])

                    report_rows = []
                    for v in self.venues:
                        used = usage_map.get(v["room_code"], 0)
                        free = max(0, v["capacity"] - used)
                        report_rows.append({
                            "Room No.": v["room_code"],
                            "Exam Capacity": v["capacity"],
                            "Block": v["building"],
                            "Alloted": used,
                            "Vacant (B-C)": free,
                        })

                    # Safe sheet name
                    s_name = f"{str(d).split()[0].replace('-', '_')}_{s}"
                    for bad in '[]:*?/\\':
                        s_name = s_name.replace(bad, "_")
                    
                    pd.DataFrame(report_rows).to_excel(writer, sheet_name=s_name[:31], index=False)

            self.log.info("Reports saved to %s and %s", out_1, out_2)

        except Exception as e:
            self.log.exception("Report generation failed: %s", e)
            raise

    def create_attendance_files(self, img_source, default_icon, pdf_root=None):
        """Generate PDF attendance sheets for every room/subject combination."""
        if pdf_root is None:
            pdf_root = os.path.join(self.result_dir, "attendance")
        os.makedirs(pdf_root, exist_ok=True)
        
        self.log.info("Starting PDF generation in: %s", pdf_root)

        # Re-organize data for printing
        print_queue = {}
        for _, alloc_list in self.final_assignments.items():
            for item in alloc_list:
                key = (str(item["date"]), str(item["slot"]), str(item["room"]), str(item["subject"]))
                print_queue.setdefault(key, []).extend(item["rolls"])

        for (dt, sl, rm, sb), roll_list in print_queue.items():
            unique_rolls = list(dict.fromkeys(roll_list))
            
            # Sanitize filename
            d_simple = str(dt).split()[0].replace("-", "_").replace("/", "_")
            f_name = f"{d_simple}_{sl}_{rm}_{sb}.pdf"
            for bad in '<>:"/\\|?* ':
                f_name = f_name.replace(bad, "_")
            
            full_path = os.path.join(pdf_root, f_name)
            
            try:
                build_attendance_pdf(
                    out_path=full_path,
                    date_str=str(dt).split()[0],
                    shift=sl,
                    room_no=rm,
                    subject_code=sb,
                    subject_name=sb,
                    roll_list=unique_rolls,
                    roll_to_name=self.student_names,
                    photos_dir=img_source,
                    no_image_icon=default_icon,
                    logger=self.log,
                )
            except Exception:
                self.log.error(f"Failed to create PDF for {sb} in {rm}")
                continue

        self.log.info("PDF generation complete.")