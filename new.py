# --- main_app.py ---

import os
import tempfile
import shutil
import requests
from flask import Flask, render_template, request, send_from_directory, flash, redirect, url_for, session, jsonify
from pathlib import Path
from werkzeug.utils import secure_filename
from threading import Thread
import uuid

# Attempt to import PDF libraries
try:
    from PyPDF2 import PdfMerger
except ImportError:
    print("ERROR: PyPDF2 not found. Please install it: pip install pypdf2")
    PdfMerger = None

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("ERROR: pypdf not found. Please install it: pip install pypdf")
    PdfReader = None
    PdfWriter = None

# --- Configuration ---
app = Flask(__name__)
app.config['SECRET_KEY'] = os.urandom(24)
app.config['GENERATED_FILE_DIR'] = os.path.join(tempfile.gettempdir(), 'past_paper_generator')
os.makedirs(app.config['GENERATED_FILE_DIR'], exist_ok=True)

tasks_status = {}

SUBJECT_NAMES = {
    "9700": "biology",
    "9701": "chemistry",
    "9702": "physics",
    "9709": "mathematics",
    "9231": "mathematics-further",
    "9618": "computer-science",
    "9608": "computer-science",
    "9626": "information-technology",
    "9084": "law",
    "9189": "history",
    "9990": "psychology",
    "9699": "sociology",
    "9696": "geography",
    "9093": "english-language",
    "9695": "english-literature",
    "9708": "economics",
    "9609": "business",
}

ALL_KEYWORD_MAPS = {
    "9231": {
        "roots": "Polynomial Roots", "cubic equation": "Polynomial Roots", "sum of roots": "Polynomial Roots", "product of roots": "Polynomial Roots",
        "vector": "Vectors", "plane": "Vectors", "line": "Vectors", "intersection": "Vectors", "perpendicular distance": "Vectors", "scalar product": "Vectors",
        "mathematical induction": "Mathematical Induction", "induction": "Mathematical Induction",
        "matrix": "Matrices", "matrices": "Matrices", "inverse of a": "Matrices", "invariant line": "Matrices", "transformation": "Matrices", "enlargement": "Matrices", "shear": "Matrices",
        "method of differences": "Series Summation", "summation": "Series Summation", "series": "Series Summation",
        "polar equation": "Polar Coordinates", "polar coordinates": "Polar Coordinates",
        "asymptotes": "Curve Sketching", "sketch": "Curve Sketching", "stationary points": "Curve Sketching",
        "Cartesian equation":"Polar Coordinates","initial line":"Polar Coordinates","pole": "Polar Coordinates",
        "elastic string": "Hooke's Law",
        "elastic spring": "Hooke's Law",
        "modulus of elasticity": "Hooke's Law",
        "natural length": "Hooke's Law",
        "extension": "Hooke's Law",
        "horizontal circle": "Circular Motion",
        "angular speed": "Circular Motion",
        "circular path": "Circular Motion",
        "vertical circle": "Circular Motion",
        "peg": "Circular Motion ",
        "string taut": "Circular Motion",
        "string becomes slack": "Circular Motion ",
        "collide": "Collisions / Impulse and Momentum",
        "collision": "Collisions / Impulse and Momentum",
        "spheres": "Collisions / Impulse and Momentum",
        "coefficient of restitution": "Collisions / Impulse and Momentum",
        "line of centres": "Collisions / Impulse and Momentum",
        "impulse": "Collisions / Impulse and Momentum",
        "momentum": "Collisions / Impulse and Momentum",
        "uniform rod": "Statics of Rigid Bodies / Equilibrium",
        "equilibrium": "Statics of Rigid Bodies / Equilibrium",
        "rough wall": "Statics of Rigid Bodies / Equilibrium",
        "friction": "Statics of Rigid Bodies / Equilibrium",
        "coefficient of friction": "Statics of Rigid Bodies / Equilibrium",
        "about to slip": "Statics of Rigid Bodies / Equilibrium",
        "frictional force": "Statics of Rigid Bodies / Equilibrium",
        "moments": "Statics of Rigid Bodies / Equilibrium",
        "projected": "Projectile Motion",
        "trajectory": "Projectile Motion",
        "horizontal ground": "Projectile Motion",
        "range": "Projectile Motion",
        "time of flight": "Projectile Motion",
        "resistive force": "Variable Force Motion ",
        "resistance": "Variable Force Motion ",
        "terminal velocity": "Variable Force Motion",
    },
    "9618": {
        "bitmapped image": "Information representation",
        "vector graphic": "Information representation",
        "colour depth": "Information representation",
        "image resolution": "Information representation",
        "file size": "Information representation",
        "file header": "Information representation",
        "lossless compression": "Information representation",
        "binary": "Information representation",
        "denary": "Information representation",
        "hexadecimal": "Information representation",
        "bit manipulation": "Information representation",
        "bitwise operation": "Information representation",
        "logical shift": "Information representation",
        "email": "Internet and Communication",
        "internet": "Internet and Communication",
        "protocol": "Internet and Communication",
        "IP address": "Internet and Communication",
        "MAC address": "Internet and Communication",
        "embedded system": "Hardware",
        "RAM": "Hardware",
        "DRAM": "Hardware",
        "SRAM": "Hardware",
        "ROM": "Hardware",
        "EPROM": "Hardware",
        "EEPROM": "Hardware",
        "monitor": "Hardware",
        "screen resolution": "Hardware",
        "pixels": "Hardware",
        "VGA port": "Hardware",
        "HDMI port": "Hardware",
        "speakers": "Hardware",
        "sensor": "Hardware",
        "input device": "Hardware",
        "output device": "Hardware",
        "storage device": "Hardware",
        "logic gate": "Processor Fundamentals",
        "logic circuit": "Processor Fundamentals",
        "truth table": "Processor Fundamentals",
        "Boolean expression": "Processor Fundamentals",
        "Von Neumann model": "Processor Fundamentals",
        "processor": "Processor Fundamentals",
        "CPU": "Processor Fundamentals",
        "register": "Processor Fundamentals",
        "Program Counter (PC)": "Processor Fundamentals",
        "Memory Address Register (MAR)": "Processor Fundamentals",
        "Memory Data Register (MDR)": "Processor Fundamentals",
        "Accumulator (ACC)": "Processor Fundamentals",
        "Index Register (IX)": "Processor Fundamentals",
        "instruction set": "Processor Fundamentals",
        "assembly language": "Processor Fundamentals",
        "opcode": "Processor Fundamentals",
        "operand": "Processor Fundamentals",
        "addressing mode": "Processor Fundamentals",
        "fetch-decode-execute cycle": "Processor Fundamentals",
        "Operating System (OS)": "System Software",
        "process management": "System Software",
        "memory management": "System Software",
        "file management": "System Software",
        "interrupt": "System Software",
        "Integrated Development Environment (IDE)": "System Software",
        "debugger": "System Software",
        "translator": "System Software",
        "compiler": "System Software",
        "interpreter": "System Software",
        "assembler": "System Software",
        "program library": "System Software",
        "utility software": "System Software",
        "security": "Security, privacy and data integrity",
        "privacy": "Security, privacy and data integrity",
        "data integrity": "Security, privacy and data integrity",
        "unauthorised access": "Security, privacy and data integrity",
        "malware": "Security, privacy and data integrity",
        "virus": "Security, privacy and data integrity",
        "hacking": "Security, privacy and data integrity",
        "firewall": "Security, privacy and data integrity",
        "encryption": "Security, privacy and data integrity",
        "authentication": "Security, privacy and data integrity",
        "password": "Security, privacy and data integrity",
        "data validation": "Security, privacy and data integrity",
        "data verification": "Security, privacy and data integrity",
        "ethics": "Ethics and Ownership",
        "ethical": "Ethics and Ownership",
        "copyright": "Ethics and Ownership",
        "plagiarism": "Ethics and Ownership",
        "software licence": "Ethics and Ownership",
        "freeware": "Ethics and Ownership",
        "shareware": "Ethics and Ownership",
        "open source": "Ethics and Ownership",
        "Data Protection Act": "Ethics and Ownership",
        "database": "Databases",
        "relational database": "Databases",
        "DML":"Databases",
        "DDL":"Databases",
        "field": "Databases",
        "primary key": "Databases",
        "foreign key": "Databases",
        "relationship": "Databases",
        "normalisation": "Databases",
        "SQL": "Databases",
        "query": "Databases",
        "SELECT statement": "Databases",
        "FROM clause": "Databases",
        "WHERE clause": "Databases",
        "data dictionary": "Databases",
        "DBMS": "Databases",
    },
    "9608": {
        "bitmapped image": "Information representation",
        "vector graphic": "Information representation",
        "colour depth": "Information representation",
        "image resolution": "Information representation",
        "file size": "Information representation",
        "file header": "Information representation",
        "lossless compression": "Information representation",
        "binary": "Information representation",
        "denary": "Information representation",
        "hexadecimal": "Information representation",
        "bit manipulation": "Information representation",
        "bitwise operation": "Information representation",
        "logical shift": "Information representation",
        "email": "Internet and Communication",
        "internet": "Internet and Communication",
        "protocol": "Internet and Communication",
        "IP address": "Internet and Communication",
        "MAC address": "Internet and Communication",
        "embedded system": "Hardware",
        "RAM": "Hardware",
        "DRAM": "Hardware",
        "SRAM": "Hardware",
        "ROM": "Hardware",
        "EPROM": "Hardware",
        "EEPROM": "Hardware",
        "monitor": "Hardware",
        "screen resolution": "Hardware",
        "VGA port": "Hardware",
        "HDMI port": "Hardware",
        "speakers": "Hardware",
        "sensor": "Hardware",
        "input device": "Hardware",
        "output device": "Hardware",
        "storage device": "Hardware",
        "logic gate": "Processor Fundamentals",
        "logic circuit": "Processor Fundamentals",
        "truth table": "Processor Fundamentals",
        "Boolean expression": "Processor Fundamentals",
        "Von Neumann model": "Processor Fundamentals",
        "processor": "Processor Fundamentals",
        "CPU": "Processor Fundamentals",
        "register": "Processor Fundamentals",
        "Program Counter (PC)": "Processor Fundamentals",
        "Memory Address Register (MAR)": "Processor Fundamentals",
        "Memory Data Register (MDR)": "Processor Fundamentals",
        "Accumulator (ACC)": "Processor Fundamentals",
        "Index Register (IX)": "Processor Fundamentals",
        "instruction set": "Processor Fundamentals",
        "assembly language": "Processor Fundamentals",
        "opcode": "Processor Fundamentals",
        "operand": "Processor Fundamentals",
        "addressing mode": "Processor Fundamentals",
        "fetch-decode-execute cycle": "Processor Fundamentals",
        "Operating System (OS)": "System Software",
        "process management": "System Software",
        "memory management": "System Software",
        "file management": "System Software",
        "interrupt": "System Software",
        "Integrated Development Environment (IDE)": "System Software",
        "debugger": "System Software",
        "translator": "System Software",
        "compiler": "System Software",
        "interpreter": "System Software",
        "assembler": "System Software",
        "program library": "System Software",
        "utility software": "System Software",
        "security": "Security, privacy and data integrity",
        "privacy": "Security, privacy and data integrity",
        "data integrity": "Security, privacy and data integrity",
        "unauthorised access": "Security, privacy and data integrity",
        "malware": "Security, privacy and data integrity",
        "virus": "Security, privacy and data integrity",
        "hacking": "Security, privacy and data integrity",
        "firewall": "Security, privacy and data integrity",
        "encryption": "Security, privacy and data integrity",
        "authentication": "Security, privacy and data integrity",
        "password": "Security, privacy and data integrity",
        "data validation": "Security, privacy and data integrity",
        "data verification": "Security, privacy and data integrity",
        "ethics": "Ethics and Ownership",
        "ethical": "Ethics and Ownership",
        "copyright": "Ethics and Ownership",
        "plagiarism": "Ethics and Ownership",
        "software licence": "Ethics and Ownership",
        "freeware": "Ethics and Ownership",
        "shareware": "Ethics and Ownership",
        "open source": "Ethics and Ownership",
        "Data Protection Act": "Ethics and Ownership",
        "database": "Databases",
        "relational database": "Databases",
        "table": "Databases",
        "record": "Databases",
        "field": "Databases",
        "primary key": "Databases",
        "foreign key": "Databases",
        "relationship": "Databases",
        "normalisation": "Databases",
        "SQL": "Databases",
        "query": "Databases",
        "SELECT statement": "Databases",
        "FROM clause": "Databases",
        "WHERE clause": "Databases",
        "data dictionary": "Databases",
        "DBMS": "Databases",
    },
    "9701": {
        "enthalpy change of reaction": "Chemical Energetics",
        "standard enthalpy change": "Chemical Energetics",
        "enthalpy change of formation": "Chemical Energetics",
        "enthalpy change of combustion": "Chemical Energetics",
        "enthalpy change of neutralization": "Chemical Energetics",
        "hess's Law": "Chemical Energetics",
        "bond enthalpy": "Chemical Energetics",
        "enthalpy change of solution": "Chemical Energetics",
        "enthalpy change of hydration": "Chemical Energetics",
        "exothermic reaction": "Chemical Energetics",
        "endothermic reaction": "Chemical Energetics",
        "state symbols": "Chemical Energetics",
        "calorimetry": "Chemical Energetics",
        "specific heat capacity": "Chemical Energetics",
        "thermal stability": "Chemical Energetics",
        "Group 2": "Chemical Energetics",
        "average bond enthalpy": "Chemical Energetics",
        "bond energies": "Chemical Energetics",
        "activation energy": "Chemical Energetics",
        "hydration": "Chemical Energetics",
        "Born-Haber" : "Chemical Energetics",
        "electron affinity": "Chemical Energetics",
        "lattice energy": "Chemical Energetics",
        "Electrode Potential": "Electrochemistry",
        "cell": "Electrochemistry",
        "half-cell": "Electrochemistry",
        "electrochemical series": "Electrochemistry",
        "voltaic cell": "Electrochemistry",
        "electrolytic cell": "Electrochemistry",
        "Nernst equation": "Electrochemistry",
        "electrolytic cell": "Electrochemistry",
        "electrode": "Electrochemistry",
        "electrolyte": "Electrochemistry",
        "overpotential": "Electrochemistry",
        "selective discharge": "Electrochemistry",
        "electrolysed": "Electrochemistry",
        "quantitative electrolysis": "Electrochemistry",
        "redox titration": "Electrochemistry",
        "balancing redox equations": "Electrochemistry",
        "oxidation number": "Electrochemistry",
        "Batteries": "Electrochemistry",
        "Fuel cells": "Electrochemistry",
        "standard electrode potential": "Electrochemistry",
        "electroplating": "Electrochemistry",
        "cell reaction": "Electrochemistry",
    }
}

def run_download_and_merge(task_id, subject_code, paper_number, start_year, end_year, sessions, include_ms):
    tasks_status[task_id] = {'status': 'Processing', 'progress': 'Starting download/merge...', 'files': {}, 'errors': []}

    if not PdfMerger:
        tasks_status[task_id]['status'] = 'Error'
        tasks_status[task_id]['errors'].append("PDF Merging library (PyPDF2) not available.")
        return {}

    base_dir = Path(app.config['GENERATED_FILE_DIR']) / task_id
    try:
        base_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        tasks_status[task_id]['status'] = 'Error'
        tasks_status[task_id]['errors'].append(f"Error creating temporary directory: {e}")
        return {}

    subject_name = SUBJECT_NAMES.get(subject_code)
    if not subject_name:
        tasks_status[task_id]['status'] = 'Error'
        tasks_status[task_id]['errors'].append(f"Subject URL name not found for code {subject_code}. Check SUBJECT_NAMES.")
        return {}

    years = list(range(start_year, end_year + 1))
    base_url = "https://bestexamhelp.com/exam/cambridge-international-a-level"

    results = {'qp': None, 'ms': None}
    file_types_to_process = ['qp']
    if include_ms:
        file_types_to_process.append('ms')

    total_files_to_attempt = 0
    for file_type in file_types_to_process:
        for year in years:
            for session in sessions:
                variants_count = 1 if session == "m" else 3
                total_files_to_attempt += variants_count

    files_processed_count = 0
    total_steps = total_files_to_attempt * (2 if include_ms else 1)

    for file_type in file_types_to_process:
        pdf_files_to_merge = []
        tasks_status[task_id]['progress'] = f'Processing {file_type.upper()}...'
        print(f"Task {task_id}: Processing {file_type.upper()}...")

        for year in years:
            for session in sessions:
                variant_numbers = ["2"] if session == "m" else ["1", "2", "3"]
                variants_full = [f"{paper_number}{v}" for v in variant_numbers]

                for variant in variants_full:
                    files_processed_count += 1
                    progress_percent = int((files_processed_count / max(1, total_steps)) * 100)
                    tasks_status[task_id]['progress'] = f'Downloading {file_type.upper()} {year}/{session}/var{variant[-1]}... ({progress_percent}%)'

                    filename = f"{subject_code}_{session}{str(year)[-2:]}_{file_type}_{variant}.pdf"
                    file_url = f"{base_url}/{subject_name}-{subject_code}/{year}/{filename}"
                    file_filepath = base_dir / filename

                    try:
                        print(f"Task {task_id}: Attempting download: {file_url}")
                        response = requests.get(file_url, timeout=30)
                        response.raise_for_status()

                        with open(file_filepath, "wb") as f:
                            f.write(response.content)
                        pdf_files_to_merge.append(file_filepath)
                        print(f"Task {task_id}: Downloaded: {filename}")

                    except requests.exceptions.HTTPError as e:
                        error_msg = f"Download failed for {filename} (HTTP {e.response.status_code}): URL={file_url}"
                        tasks_status[task_id]['errors'].append(error_msg)
                        print(f"Task {task_id}: {error_msg}")
                    except requests.exceptions.RequestException as e:
                        error_msg = f"Network error downloading {filename}: {e}"
                        tasks_status[task_id]['errors'].append(error_msg)
                        print(f"Task {task_id}: {error_msg}")
                    except Exception as e:
                        error_msg = f"Unexpected error downloading {filename}: {e}"
                        tasks_status[task_id]['errors'].append(error_msg)
                        print(f"Task {task_id}: {error_msg}")

        if pdf_files_to_merge:
            tasks_status[task_id]['progress'] = f'Merging {len(pdf_files_to_merge)} {file_type.upper()} files... ({progress_percent}%)'
            output_filename = f"{subject_code}_{paper_number}_{start_year}-{end_year}_{''.join(sessions)}_{file_type}_merged.pdf"
            output_filepath = base_dir / output_filename
            try:
                merger = PdfMerger()
                for pdf_file in pdf_files_to_merge:
                    try:
                        merger.append(str(pdf_file))
                    except Exception as merge_err:
                        error_msg = f"Could not append file {pdf_file.name} to {file_type.upper()} merge: {merge_err}. Skipping."
                        tasks_status[task_id]['errors'].append(error_msg)
                        print(f"Task {task_id}: {error_msg}")

                if len(merger.pages) > 0:
                    merger.write(str(output_filepath))
                    merger.close()
                    results[file_type] = {'path': str(output_filepath), 'filename': output_filename}
                    print(f"Task {task_id}: Successfully merged {file_type.upper()} to {output_filename}")
                else:
                    error_msg = f"No valid pages found/appended to merge for {file_type.upper()}."
                    tasks_status[task_id]['errors'].append(error_msg)
                    print(f"Task {task_id}: {error_msg}")
                    if file_type in results: del results[file_type]

            except Exception as e:
                error_msg = f"Error during merging process for {file_type.upper()} PDF files: {e}"
                tasks_status[task_id]['errors'].append(error_msg)
                print(f"Task {task_id}: {error_msg}")
                if file_type in results: del results[file_type]
        else:
            error_msg = f"No {file_type.upper()} files were successfully downloaded/found to merge."
            tasks_status[task_id]['errors'].append(error_msg)
            print(f"Task {task_id}: {error_msg}")

    tasks_status[task_id]['progress'] = 'Download/Merge phase complete. Checking for topical generation.'
    tasks_status[task_id]['files'].update(results)
    return results

def run_create_topical(task_id, merged_qp_path_str, subject_code):
    tasks_status[task_id]['progress'] = 'Starting topical generation...'
    print(f"Task {task_id}: Starting topical generation for subject {subject_code}")

    if not PdfReader or not PdfWriter:
        error_msg = "PDF processing library (pypdf) not available."
        tasks_status[task_id]['errors'].append(error_msg)
        tasks_status[task_id]['progress'] = f'Error: {error_msg}'
        return {'topical_files': []}

    keyword_map = ALL_KEYWORD_MAPS.get(subject_code)
    if not keyword_map:
        info_msg = f"No keyword map available for subject {subject_code}. Skipping topical generation."
        tasks_status[task_id]['errors'].append(info_msg)
        tasks_status[task_id]['progress'] = info_msg
        print(f"Task {task_id}: {info_msg}")
        return {'topical_files': []}

    input_pdf_path = Path(merged_qp_path_str)
    output_dir = input_pdf_path.parent / "topical"

    if not input_pdf_path.is_file():
        error_msg = f"Merged QP PDF not found at '{input_pdf_path}' for topical generation."
        tasks_status[task_id]['errors'].append(error_msg)
        tasks_status[task_id]['progress'] = 'Error: Input for topical not found.'
        print(f"Task {task_id}: {error_msg}")
        return {'topical_files': []}

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"Task {task_id}: Topical output directory: {output_dir.resolve()}")
    except OSError as e:
        error_msg = f"Error creating topical output directory '{output_dir}': {e}"
        tasks_status[task_id]['errors'].append(error_msg)
        tasks_status[task_id]['progress'] = 'Error creating topical directory.'
        print(f"Task {task_id}: {error_msg}")
        return {'topical_files': []}

    topic_pages = {topic: set() for topic in keyword_map.values()}
    topical_files_generated = []
    pages_processed = 0
    num_pages = 0

    try:
        reader = PdfReader(input_pdf_path)
        num_pages = len(reader.pages)
        tasks_status[task_id]['progress'] = f'Analyzing {num_pages} pages for topics...'
        print(f"Task {task_id}: Analyzing {num_pages} pages in {input_pdf_path.name}")

        for i, page in enumerate(reader.pages):
            try:
                text = page.extract_text()
                if text:
                    text_lower = text.lower()
                    page_matched_this_loop = False
                    for keyword, topic in keyword_map.items():
                        if keyword in text_lower:
                            topic_pages[topic].add(i)
                            page_matched_this_loop = True

                pages_processed += 1
                if i > 0 and i % 50 == 0:
                    progress_percent = int((pages_processed/max(1, num_pages))*100)
                    tasks_status[task_id]['progress'] = f'Analyzing pages... ({progress_percent}%)'

            except Exception as e:
                tasks_status[task_id]['errors'].append(f"Warning: Error extracting text from page {i+1}: {e}.")
                print(f"Task {task_id}: Warning - Error extracting text from page {i+1}: {e}")

        tasks_status[task_id]['progress'] = f'Page analysis complete. Found potential matches for {len([t for t, p in topic_pages.items() if p])} topics.'
        print(f"Task {task_id}: Page analysis complete.")

        topics_created_count = 0
        total_topics_with_pages = len([t for t, p in topic_pages.items() if p])

        if total_topics_with_pages == 0:
            info_msg = f"No keywords from the map for {subject_code} were matched in the merged PDF. No topical files generated."
            tasks_status[task_id]['errors'].append(info_msg)
            print(f"Task {task_id}: {info_msg}")
        else:
            print(f"Task {task_id}: Found pages for {total_topics_with_pages} topics. Starting PDF creation...")

        for topic, page_indices in topic_pages.items():
            if page_indices:
                sorted_indices = sorted(list(page_indices))
                writer = PdfWriter()
                progress_percent = int((topics_created_count / max(1, total_topics_with_pages)) * 100)
                tasks_status[task_id]['progress'] = f'Creating PDF for topic: {topic}... ({progress_percent}%)'
                print(f"Task {task_id}: Creating PDF for '{topic}' with {len(sorted_indices)} pages...")

                for page_index in sorted_indices:
                    if 0 <= page_index < num_pages:
                        try:
                            writer.add_page(reader.pages[page_index])
                        except Exception as e:
                            tasks_status[task_id]['errors'].append(f"Error adding page {page_index+1} to topic '{topic}': {e}")
                            print(f"Task {task_id}: Error adding page {page_index+1} to topic '{topic}': {e}")
                    else:
                        tasks_status[task_id]['errors'].append(f"Warning: Invalid page index {page_index+1} encountered for topic '{topic}'. Skipping.")
                        print(f"Task {task_id}: Warning - Invalid page index {page_index+1} for topic '{topic}'.")

                if len(writer.pages) > 0:
                    safe_topic_name = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in topic).strip().replace(" ", "_")
                    while "__" in safe_topic_name:
                        safe_topic_name = safe_topic_name.replace("__", "_")
                    output_pdf_filename = f"Topical_{subject_code}_{safe_topic_name}.pdf"
                    output_pdf_path = output_dir / output_pdf_filename

                    try:
                        with open(output_pdf_path, "wb") as output_file:
                            writer.write(output_file)
                        topical_files_generated.append({'path': str(output_pdf_path), 'filename': output_pdf_filename, 'topic': topic})
                        topics_created_count += 1
                        print(f"Task {task_id}: Created topical PDF: {output_pdf_filename}")
                    except Exception as e:
                        error_msg = f"Error writing PDF file for topic '{topic}': {e}"
                        tasks_status[task_id]['errors'].append(error_msg)
                        print(f"Task {task_id}: {error_msg}")
                else:
                    info_msg = f"No valid pages could be added for topic '{topic}' for subject {subject_code}. PDF not created."
                    tasks_status[task_id]['errors'].append(info_msg)
                    print(f"Task {task_id}: {info_msg}")

        tasks_status[task_id]['progress'] = f'Topical generation finished. {topics_created_count} files created.'
        print(f"Task {task_id}: Topical generation finished. Created {topics_created_count} files.")
        return {'topical_files': topical_files_generated}

    except FileNotFoundError:
        error_msg = f"Error: Input PDF file not found at '{input_pdf_path}' during topical generation."
        tasks_status[task_id]['errors'].append(error_msg)
        tasks_status[task_id]['progress'] = 'Error: Input file disappeared.'
        print(f"Task {task_id}: {error_msg}")
        return {'topical_files': []}
    except Exception as e:
        error_msg = f"An unexpected error occurred during topical PDF creation for {subject_code}: {e}"
        tasks_status[task_id]['errors'].append(error_msg)
        tasks_status[task_id]['progress'] = 'Error during topical generation.'
        print(f"Task {task_id}: {error_msg}")
        import traceback
        traceback.print_exc()
        return {'topical_files': []}

def background_task_runner(task_id, subject_code, paper_number, start_year, end_year, sessions, include_ms, generate_topical):
    try:
        merge_results = run_download_and_merge(
            task_id, subject_code, paper_number, start_year, end_year, sessions, include_ms
        )

        if tasks_status[task_id].get('status') == 'Error':
            print(f"Task {task_id}: Halting after critical merge error.")
            return

        merged_qp_info = tasks_status[task_id]['files'].get('qp')

        topical_results = {'topical_files': []}
        if generate_topical:
            if merged_qp_info and merged_qp_info.get('path'):
                merged_qp_path = merged_qp_info['path']
                topical_results = run_create_topical(task_id, merged_qp_path, subject_code)
            else:
                info_msg = "Topical generation skipped: Merged Question Paper file was not successfully created or found."
                tasks_status[task_id]['errors'].append(info_msg)
                tasks_status[task_id]['progress'] = 'Topical generation skipped (No QP).'
                print(f"Task {task_id}: {info_msg}")
        else:
            tasks_status[task_id]['progress'] = 'Topical generation not requested.'
            print(f"Task {task_id}: Topical generation not requested.")

        tasks_status[task_id]['files']['topical'] = topical_results.get('topical_files', [])

        if tasks_status[task_id].get('status') != 'Error':
            tasks_status[task_id]['status'] = 'Completed'
            tasks_status[task_id]['progress'] = 'All tasks finished.'
            print(f"Task {task_id}: Completed successfully.")

    except Exception as e:
        error_msg = f"Critical error in background task runner {task_id}: {e}"
        print(error_msg)
        import traceback
        traceback.print_exc()
        tasks_status[task_id]['status'] = 'Error'
        tasks_status[task_id]['errors'].append(error_msg)
        tasks_status[task_id]['progress'] = 'Critical error encountered.'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        topical_available_subjects = {
            code: SUBJECT_NAMES.get(code, f"Unknown Subject {code}")
            for code in ALL_KEYWORD_MAPS.keys() if code in SUBJECT_NAMES
        }
        return render_template('index.html',
                               subjects=SUBJECT_NAMES,
                               topical_available_subjects=topical_available_subjects)

    if request.method == 'POST':
        try:
            subject_code = request.form.get('subject_code')
            paper_number = request.form.get('paper_number')
            year_range = request.form.get('year_range')
            session_input = request.form.getlist('sessions')
            include_ms = 'include_ms' in request.form
            generate_topical = 'generate_topical' in request.form

            errors = []
            if not subject_code: errors.append("Please select a subject.")
            if not paper_number: errors.append("Please enter a paper number.")
            if not year_range: errors.append("Please enter a year range.")
            if not session_input: errors.append("Please select at least one session.")

            if subject_code and subject_code not in SUBJECT_NAMES:
                errors.append(f"Invalid subject code selected: {subject_code}.")

            if paper_number and (not paper_number.isdigit() or not (1 <= int(paper_number) <= 9)):
                errors.append("Invalid paper number. Please enter a single digit (e.g., 1, 2, 4).")

            start_year, end_year = None, None
            if year_range:
                try:
                    start_year_str, end_year_str = year_range.split('-')
                    start_year = int(start_year_str)
                    end_year = int(end_year_str)
                    if start_year > end_year or start_year < 2010 or end_year > 2025:
                        raise ValueError("Year range out of bounds or invalid order.")
                except (ValueError, IndexError):
                    errors.append("Invalid year range format. Use YYYY-YYYY (e.g., 2018-2023) within reasonable limits.")

            valid_sessions = {'m', 's', 'w'}
            sessions = [s for s in session_input if s in valid_sessions]
            if session_input and not sessions:
                errors.append("Invalid session selection.")

            if generate_topical and subject_code and subject_code not in ALL_KEYWORD_MAPS:
                flash(f"Warning: Topical keyword map not currently available for {SUBJECT_NAMES.get(subject_code, subject_code)}. Topical files will not be generated.", "warning")
                generate_topical = False

            if errors:
                for error in errors:
                    flash(error, "danger")
                topical_available_subjects = { code: SUBJECT_NAMES.get(code, f"Unknown Subject {code}") for code in ALL_KEYWORD_MAPS.keys() if code in SUBJECT_NAMES }
                return render_template('index.html',
                                        subjects=SUBJECT_NAMES,
                                        topical_available_subjects=topical_available_subjects,
                                        selected_subject=subject_code,
                                        entered_paper=paper_number,
                                        entered_years=year_range,
                                        selected_sessions=session_input,
                                        checked_ms=include_ms,
                                        checked_topical=generate_topical and subject_code in ALL_KEYWORD_MAPS
                                        )

            task_id = str(uuid.uuid4())
            tasks_status[task_id] = {'status': 'Queued', 'progress': 'Initializing...', 'files': {}, 'errors': [], 'params': {
                'subject': subject_code, 'paper': paper_number, 'years': year_range, 'sessions': sessions, 'ms': include_ms, 'topical': generate_topical
            }}

            thread = Thread(target=background_task_runner, args=(
                task_id, subject_code, paper_number, start_year, end_year, sessions, include_ms, generate_topical
            ))
            thread.daemon = True
            thread.start()
            print(f"Task {task_id}: Started background thread for {subject_code} {paper_number} {year_range}")

            return redirect(url_for('task_status', task_id=task_id))

        except Exception as e:
            flash(f"An unexpected error occurred processing your request: {e}", "danger")
            print(f"Error in POST handler: {e}")
            import traceback
            traceback.print_exc()
            return redirect(url_for('index'))

@app.route('/status/<task_id>')
def task_status(task_id):
    status_info = tasks_status.get(task_id)
    if not status_info:
        flash(f"Task ID '{task_id}' not found.", "warning")
        return redirect(url_for('index'))
    return render_template('status.html', status=status_info, task_id=task_id)

@app.route('/status_api/<task_id>')
def task_status_api(task_id):
    status_info = tasks_status.get(task_id)
    if not status_info:
        return jsonify({"status": "Error", "message": "Task not found"}), 404
    return jsonify(status_info)

@app.route('/download/<task_id>/<file_key>')
def download_file(task_id, file_key):
    status_info = tasks_status.get(task_id)
    if not status_info or status_info.get('status') == 'Queued':
        flash("Task not found or not started processing.", "danger")
        return redirect(url_for('task_status', task_id=task_id))

    file_info = None
    base_directory = Path(app.config['GENERATED_FILE_DIR']) / task_id
    file_path_to_serve = None
    filename_to_serve = None

    if file_key == 'merged_qp' and status_info.get('files', {}).get('qp'):
        file_info = status_info['files']['qp']
        if file_info and 'path' in file_info and 'filename' in file_info:
            file_path_to_serve = Path(file_info['path'])
            filename_to_serve = file_info['filename']
            if not str(file_path_to_serve.parent) == str(base_directory):
                print(f"Security warning: Attempt to download file outside task dir: {file_path_to_serve}")
                file_path_to_serve = None

    elif file_key == 'merged_ms' and status_info.get('files', {}).get('ms'):
        file_info = status_info['files']['ms']
        if file_info and 'path' in file_info and 'filename' in file_info:
            file_path_to_serve = Path(file_info['path'])
            filename_to_serve = file_info['filename']
            if not str(file_path_to_serve.parent) == str(base_directory):
                print(f"Security warning: Attempt to download file outside task dir: {file_path_to_serve}")
                file_path_to_serve = None

    elif file_key.startswith('topical_'):
        try:
            index = int(file_key.split('_')[1])
            topical_files = status_info.get('files', {}).get('topical', [])
            if 0 <= index < len(topical_files):
                file_info = topical_files[index]
                if file_info and 'path' in file_info and 'filename' in file_info:
                    file_path_to_serve = Path(file_info['path'])
                    filename_to_serve = file_info['filename']
                    expected_parent = base_directory / 'topical'
                    if not str(file_path_to_serve.parent) == str(expected_parent):
                        print(f"Security warning: Attempt to download topical file outside task/topical dir: {file_path_to_serve}")
                        file_path_to_serve = None

        except (ValueError, IndexError):
            pass

    if file_path_to_serve and filename_to_serve and file_path_to_serve.is_file():
        try:
            print(f"Task {task_id}: Serving file '{filename_to_serve}' from directory '{file_path_to_serve.parent}'")
            return send_from_directory(
                directory=file_path_to_serve.parent,
                path=filename_to_serve,
                as_attachment=True
            )
        except Exception as e:
            flash(f"Error serving file '{filename_to_serve}': {e}", "danger")
            print(f"Download error for {task_id}/{file_key}: {e}")
            return redirect(url_for('task_status', task_id=task_id))
    else:
        flash(f"Requested file ({file_key}) not found or invalid for task {task_id}.", "danger")
        print(f"Task {task_id}: File download failed - File key '{file_key}', Path '{file_path_to_serve}', Filename '{filename_to_serve}'")
        return redirect(url_for('task_status', task_id=task_id))

@app.route('/cleanup/<task_id>', methods=['POST'])
def cleanup_task(task_id):
    task_dir = Path(app.config['GENERATED_FILE_DIR']) / task_id
    task_exists_in_status = task_id in tasks_status

    if task_exists_in_status:
        try:
            del tasks_status[task_id]
        except KeyError: pass

    if task_dir.exists() and task_dir.is_dir():
        try:
            shutil.rmtree(task_dir)
            flash(f"Cleaned up files for task {task_id}.", "success")
            print(f"Task {task_id}: Cleaned up directory {task_dir}")
        except OSError as e:
            flash(f"Error cleaning up files for task {task_id}: {e}", "danger")
            print(f"Cleanup error for {task_id}: {e}")
    elif task_exists_in_status:
        flash(f"Task {task_id} removed from tracking. No files/directory found to clean up.", "info")
    else:
        flash(f"Task {task_id} not found for cleanup.", "warning")

    return redirect(url_for('index'))

if __name__ == '__main__':
    if not PdfMerger or not PdfReader or not PdfWriter:
        print("\nERROR: Required PDF libraries (PyPDF2 and pypdf) are not installed.")
        print("Please install them:")
        print("pip install pypdf2 pypdf requests Flask")
    else:
        try:
            os.makedirs(app.config['GENERATED_FILE_DIR'], exist_ok=True)
            print(f"File storage directory: {app.config['GENERATED_FILE_DIR']}")
            print("Starting Flask development server...")
            app.run(debug=True, host='0.0.0.0', port=5000)
        except Exception as e:
            print(f"Failed to start Flask app: {e}")
            import traceback
            traceback.print_exc()
