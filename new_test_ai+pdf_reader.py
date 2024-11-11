import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import json
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import requests
from typing import Dict, List, Optional
import threading
import queue
from datetime import datetime
import csv
import openai
from serpapi import GoogleSearch
import time

class AIDataExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Web Data Extractor")
        self.root.state('zoomed')
        
        # Initialize variables
        self.data_df = None
        self.selected_column = None
        self.search_results = {}
        self.extracted_data = {}
        self.processing_queue = queue.Queue()
        
        # API configurations
        self.setup_api_config()
        
        # Create main interface
        self.create_gui()
        
        # Start processing thread
        self.processing_thread = threading.Thread(target=self.process_queue, daemon=True)
        self.processing_thread.start()

    def setup_api_config(self):
        # Load API keys from config file
        try:
            with open('config.json', 'r') as f:
                config = json.load(f)
                self.serp_api_key = config.get('7207cefe6951100ec76b518685088ee70afd29b3dd459992bcfdacea44ad26d2')
                self.openai_api_key = config.get('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')
                self.google_creds_path = config.get('google_credentials_path')
        except FileNotFoundError:
            self.serp_api_key = ''
            self.openai_api_key = ''
            self.google_creds_path = ''

    def create_gui(self):
        # Create main container
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create sections
        self.create_data_input_section()
        self.create_query_section()
        self.create_results_section()
        self.create_status_bar()

    def create_data_input_section(self):
        # Data Input Frame
        input_frame = ttk.LabelFrame(self.main_container, text="Data Input")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # File Upload
        upload_frame = ttk.Frame(input_frame)
        upload_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(upload_frame, text="Upload CSV", command=self.upload_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(upload_frame, text="Connect Google Sheet", command=self.connect_google_sheet).pack(side=tk.LEFT, padx=5)
        
        # Column Selection
        column_frame = ttk.Frame(input_frame)
        column_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(column_frame, text="Select Main Column:").pack(side=tk.LEFT, padx=5)
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(column_frame, textvariable=self.column_var, state='readonly')
        self.column_combo.pack(side=tk.LEFT, padx=5)
        self.column_combo.bind('<<ComboboxSelected>>', self.on_column_select)
        
        # Preview Frame
        preview_frame = ttk.Frame(input_frame)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Preview Table
        self.preview_tree = ttk.Treeview(preview_frame, show="headings")
        preview_scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        preview_scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)
        
        preview_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        preview_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def create_query_section(self):
        # Query Frame
        query_frame = ttk.LabelFrame(self.main_container, text="Query Configuration")
        query_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Prompt Template
        prompt_frame = ttk.Frame(query_frame)
        prompt_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(prompt_frame, text="Query Template:").pack(side=tk.LEFT, padx=5)
        self.prompt_var = tk.StringVar(value="Get me the email address of {entity}")
        prompt_entry = ttk.Entry(prompt_frame, textvariable=self.prompt_var, width=50)
        prompt_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Extraction Template
        extract_frame = ttk.Frame(query_frame)
        extract_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(extract_frame, text="Extraction Template:").pack(side=tk.LEFT, padx=5)
        self.extract_var = tk.StringVar(value="Extract the email address from the following web results for {entity}")
        extract_entry = ttk.Entry(extract_frame, textvariable=self.extract_var, width=50)
        extract_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Control Buttons
        control_frame = ttk.Frame(query_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(control_frame, text="Start Processing", command=self.start_processing).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Stop Processing", command=self.stop_processing).pack(side=tk.LEFT, padx=5)

    def create_results_section(self):
        # Results Frame
        results_frame = ttk.LabelFrame(self.main_container, text="Results")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Results Table
        self.results_tree = ttk.Treeview(results_frame, show="headings")
        results_scroll_y = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        results_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=results_scroll_y.set, xscrollcommand=results_scroll_x.set)
        
        results_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        results_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Export Buttons
        export_frame = ttk.Frame(results_frame)
        export_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(export_frame, text="Download CSV", command=self.export_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Update Google Sheet", command=self.update_google_sheet).pack(side=tk.LEFT, padx=5)

    def create_status_bar(self):
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.main_container, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X)

    def upload_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                self.data_df = pd.read_csv(file_path)
                self.update_column_list()
                self.update_preview()
                self.status_var.set(f"Loaded CSV file: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {str(e)}")

    def connect_google_sheet(self):
        try:
            creds = None
            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/spreadsheets.readonly'])
            
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(self.google_creds_path, ['https://www.googleapis.com/auth/spreadsheets.readonly'])
                    creds = flow.run_local_server(port=0)
                
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())
            
            # Get Google Sheet ID from user
            sheet_id = tk.simpledialog.askstring("Input", "Enter Google Sheet ID:")
            if sheet_id:
                service = build('sheets', 'v4', credentials=creds)
                result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range='A1:ZZ').execute()
                values = result.get('values', [])
                
                if not values:
                    raise ValueError("No data found in sheet")
                
                # Convert to DataFrame
                self.data_df = pd.DataFrame(values[1:], columns=values[0])
                self.update_column_list()
                self.update_preview()
                self.status_var.set("Connected to Google Sheet successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect to Google Sheet: {str(e)}")

    def update_column_list(self):
        if self.data_df is not None:
            self.column_combo['values'] = list(self.data_df.columns)
            if len(self.data_df.columns) > 0:
                self.column_combo.set(self.data_df.columns[0])

    def update_preview(self):
        if self.data_df is not None:
            # Clear existing items
            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)
            
            # Configure columns
            self.preview_tree['columns'] = list(self.data_df.columns)
            for col in self.data_df.columns:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=100)
            
            # Add data (first 10 rows)
            for idx, row in self.data_df.head(10).iterrows():
                self.preview_tree.insert("", tk.END, values=list(row))

    def on_column_select(self, event):
        self.selected_column = self.column_var.get()

    def start_processing(self):
        if not self.selected_column or self.data_df is None:
            messagebox.showwarning("Warning", "Please select a data source and column first")
            return
        
        # Clear previous results
        self.search_results.clear()
        self.extracted_data.clear()
        
        # Configure results tree
        self.results_tree['columns'] = ['Entity', 'Extracted Data', 'Status']
        for col in ['Entity', 'Extracted Data', 'Status']:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=200)
        
        # Queue processing tasks
        for entity in self.data_df[self.selected_column]:
            self.processing_queue.put(entity)
        
        self.status_var.set("Processing started...")

    def process_queue(self):
        while True:
            try:
                entity = self.processing_queue.get(timeout=1)
                self.process_entity(entity)
                self.processing_queue.task_done()
            except queue.Empty:
                time.sleep(0.1)
            except Exception as e:
                print(f"Error processing entity: {str(e)}")

    def process_entity(self, entity):
        try:
            # Update status
            self.status_var.set(f"Processing: {entity}")
            
            # Perform web search
            search_results = self.perform_web_search(entity)
            self.search_results[entity] = search_results
            
            # Extract information using LLM
            extracted_info = self.extract_information(entity, search_results)
            self.extracted_data[entity] = extracted_info
            
            # Update results tree
            self.results_tree.insert("", tk.END, values=[entity, extracted_info, "Complete"])
            
        except Exception as e:
            self.results_tree.insert("", tk.END, values=[entity, str(e), "Error"])

    def perform_web_search(self, entity):
        try:
            query = self.prompt_var.get().format(entity=entity)
            params = {
                "api_key": self.serp_api_key,
                "engine": "google",
                "q": query,
                "num": 5
            }
            
            search = GoogleSearch(params)
            results = search.get_dict()
            
            # Extract relevant information from results
            snippets = []
            for result in results.get('organic_results', []):
                snippets.append(result.get('snippet', ''))
            
            return '\n'.join(snippets)
            
        except Exception as e:
            raise Exception(f"Search failed: {str(e)}")

    def extract_information(self, entity, search_results):
        try:
            prompt = self.extract_var.get().format(entity=entity)
            
            # Use OpenAI API for extraction
            openai.api_key = self.openai_api_key
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that extracts specific information from web search results."},
                    {"role": "user", "content": f"{prompt}\n\nSearch Results:\n{search_results}"}
                ]
            )
            
            return response.choices[0].message['content'].strip()
            
        except Exception as e:
            raise Exception(f"Extraction failed: {str(e)}")

    def stop_processing(self):
        with self.processing_queue.mutex:
            self.processing