import win32com.client
import pandas as pd
import json
import os
import time
import threading
from datetime import datetime, timedelta
from pathlib import Path
import re
from typing import Dict, List, Optional, Any
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from sentence_transformers import SentenceTransformer
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import pickle
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class OutlookAutomationTool:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.sent_items = None
        self.live_tracking = False
        self.tracking_thread = None
        self.email_embeddings = {}
        self.model = None
        self.initialize_outlook()
        self.initialize_ml_model()
        
    def initialize_outlook(self):
        """Initialize Outlook COM object"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # Inbox
            self.sent_items = self.namespace.GetDefaultFolder(5)  # Sent Items
            logger.info("Outlook connection established successfully")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
            raise
    
    def initialize_ml_model(self):
        """Initialize lightweight ML model for semantic search"""
        try:
            # Using a lightweight sentence transformer model
            self.model = SentenceTransformer('all-MiniLM-L6-v2')
            logger.info("ML model initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize ML model: {e}")
            # Fallback to basic keyword search if model fails
            self.model = None
    
    def send_template_emails(self, template_path: str, recipients_data: List[Dict], 
                           subject_template: str, attachments: List[str] = None):
        """Send emails based on template with variable replacement"""
        try:
            # Read email template
            with open(template_path, 'r', encoding='utf-8') as file:
                email_template = file.read()
            
            sent_count = 0
            failed_recipients = []
            
            for recipient_data in recipients_data:
                try:
                    # Create new mail item
                    mail = self.outlook.CreateItem(0)  # Mail item
                    
                    # Replace placeholders in template
                    email_body = email_template
                    subject = subject_template
                    
                    for key, value in recipient_data.items():
                        placeholder = f"{{{key}}}"
                        email_body = email_body.replace(placeholder, str(value))
                        subject = subject.replace(placeholder, str(value))
                    
                    # Set email properties
                    mail.To = recipient_data.get('email', '')
                    mail.Subject = subject
                    mail.Body = email_body
                    
                    # Add attachments if provided
                    if attachments:
                        for attachment_path in attachments:
                            if os.path.exists(attachment_path):
                                mail.Attachments.Add(attachment_path)
                    
                    # Send email
                    mail.Send()
                    sent_count += 1
                    logger.info(f"Email sent to {recipient_data.get('email', 'Unknown')}")
                    
                except Exception as e:
                    failed_recipients.append(recipient_data.get('email', 'Unknown'))
                    logger.error(f"Failed to send email to {recipient_data.get('email', 'Unknown')}: {e}")
            
            return {
                'sent_count': sent_count,
                'failed_recipients': failed_recipients,
                'total_recipients': len(recipients_data)
            }
            
        except Exception as e:
            logger.error(f"Error in send_template_emails: {e}")
            raise
    
    def track_email_responses(self, subject_filter: str = None, days_back: int = 7):
        """Track email responses including voting, attachments, etc."""
        try:
            # Calculate date range
            start_date = datetime.now() - timedelta(days=days_back)
            
            # Search for emails
            emails = []
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)
            
            for message in messages:
                try:
                    if message.ReceivedTime >= start_date:
                        if subject_filter and subject_filter.lower() not in message.Subject.lower():
                            continue
                        
                        email_data = {
                            'subject': message.Subject,
                            'sender': message.SenderEmailAddress,
                            'sender_name': message.SenderName,
                            'received_time': message.ReceivedTime,
                            'body': message.Body[:500] + "..." if len(message.Body) > 500 else message.Body,
                            'has_attachments': message.Attachments.Count > 0,
                            'attachment_count': message.Attachments.Count,
                            'attachment_names': [att.FileName for att in message.Attachments],
                            'is_voting_response': self.is_voting_response(message),
                            'voting_response': self.extract_voting_response(message),
                            'message_class': message.MessageClass,
                            'entry_id': message.EntryID
                        }
                        
                        emails.append(email_data)
                    
                except Exception as e:
                    logger.error(f"Error processing email: {e}")
                    continue
            
            return emails
            
        except Exception as e:
            logger.error(f"Error in track_email_responses: {e}")
            raise
    
    def is_voting_response(self, message):
        """Check if email is a voting response"""
        voting_keywords = ['vote', 'voting', 'approve', 'reject', 'yes', 'no']
        return any(keyword in message.Subject.lower() for keyword in voting_keywords)
    
    def extract_voting_response(self, message):
        """Extract voting response from email"""
        body_lower = message.Body.lower()
        
        # Common voting patterns
        if 'approve' in body_lower or 'yes' in body_lower:
            return 'Approved'
        elif 'reject' in body_lower or 'no' in body_lower:
            return 'Rejected'
        elif 'abstain' in body_lower:
            return 'Abstained'
        else:
            return 'Unknown'
    
    def download_attachments_by_subject(self, subject_filter: str, download_folder: str, 
                                      merge_excel: bool = False):
        """Download attachments from emails based on subject line"""
        try:
            if not os.path.exists(download_folder):
                os.makedirs(download_folder)
            
            downloaded_files = []
            excel_files = []
            
            messages = self.inbox.Items
            
            for message in messages:
                try:
                    if subject_filter.lower() in message.Subject.lower():
                        if message.Attachments.Count > 0:
                            for attachment in message.Attachments:
                                file_path = os.path.join(download_folder, attachment.FileName)
                                attachment.SaveAsFile(file_path)
                                downloaded_files.append(file_path)
                                
                                # Track Excel files for merging
                                if merge_excel and attachment.FileName.lower().endswith(('.xlsx', '.xls')):
                                    excel_files.append(file_path)
                                
                                logger.info(f"Downloaded: {attachment.FileName}")
                
                except Exception as e:
                    logger.error(f"Error downloading attachment: {e}")
                    continue
            
            # Merge Excel files if requested
            merged_file = None
            if merge_excel and excel_files:
                merged_file = self.merge_excel_files(excel_files, download_folder)
            
            return {
                'downloaded_files': downloaded_files,
                'excel_files': excel_files,
                'merged_file': merged_file,
                'total_downloads': len(downloaded_files)
            }
            
        except Exception as e:
            logger.error(f"Error in download_attachments_by_subject: {e}")
            raise
    
    def merge_excel_files(self, excel_files: List[str], output_folder: str):
        """Merge multiple Excel files into one"""
        try:
            all_data = []
            
            for file_path in excel_files:
                try:
                    df = pd.read_excel(file_path)
                    # Add source file column
                    df['source_file'] = os.path.basename(file_path)
                    all_data.append(df)
                except Exception as e:
                    logger.error(f"Error reading {file_path}: {e}")
                    continue
            
            if all_data:
                merged_df = pd.concat(all_data, ignore_index=True)
                merged_file_path = os.path.join(output_folder, f"merged_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                merged_df.to_excel(merged_file_path, index=False)
                logger.info(f"Merged file created: {merged_file_path}")
                return merged_file_path
            
            return None
            
        except Exception as e:
            logger.error(f"Error merging Excel files: {e}")
            return None
    
    def start_live_tracking(self, callback_function, interval_minutes: int = 5):
        """Start live tracking of inbox"""
        self.live_tracking = True
        
        def tracking_loop():
            last_check = datetime.now()
            
            while self.live_tracking:
                try:
                    # Get new emails since last check
                    messages = self.inbox.Items
                    messages.Sort("[ReceivedTime]", True)
                    
                    new_emails = []
                    for message in messages:
                        if message.ReceivedTime > last_check:
                            email_data = {
                                'subject': message.Subject,
                                'sender': message.SenderEmailAddress,
                                'received_time': message.ReceivedTime,
                                'has_attachments': message.Attachments.Count > 0
                            }
                            new_emails.append(email_data)
                    
                    if new_emails:
                        callback_function(new_emails)
                    
                    last_check = datetime.now()
                    time.sleep(interval_minutes * 60)
                    
                except Exception as e:
                    logger.error(f"Error in live tracking: {e}")
                    time.sleep(30)  # Wait before retrying
        
        self.tracking_thread = threading.Thread(target=tracking_loop, daemon=True)
        self.tracking_thread.start()
        logger.info("Live tracking started")
    
    def stop_live_tracking(self):
        """Stop live tracking"""
        self.live_tracking = False
        if self.tracking_thread:
            self.tracking_thread.join(timeout=5)
        logger.info("Live tracking stopped")
    
    def semantic_search_emails(self, query: str, max_results: int = 10, days_back: int = 30):
        """Search emails using semantic similarity"""
        try:
            if not self.model:
                # Fallback to keyword search
                return self.keyword_search_emails(query, max_results, days_back)
            
            # Get emails from specified time range
            start_date = datetime.now() - timedelta(days=days_back)
            emails = []
            
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)
            
            for message in messages:
                try:
                    if message.ReceivedTime >= start_date:
                        email_text = f"{message.Subject} {message.Body}"
                        emails.append({
                            'text': email_text,
                            'subject': message.Subject,
                            'sender': message.SenderEmailAddress,
                            'received_time': message.ReceivedTime,
                            'entry_id': message.EntryID
                        })
                except Exception as e:
                    continue
            
            if not emails:
                return []
            
            # Create embeddings for query and emails
            query_embedding = self.model.encode([query])
            email_texts = [email['text'] for email in emails]
            email_embeddings = self.model.encode(email_texts)
            
            # Calculate similarities
            similarities = cosine_similarity(query_embedding, email_embeddings)[0]
            
            # Get top results
            top_indices = np.argsort(similarities)[::-1][:max_results]
            
            results = []
            for idx in top_indices:
                if similarities[idx] > 0.1:  # Minimum similarity threshold
                    email = emails[idx]
                    email['similarity_score'] = float(similarities[idx])
                    results.append(email)
            
            return results
            
        except Exception as e:
            logger.error(f"Error in semantic search: {e}")
            return self.keyword_search_emails(query, max_results, days_back)
    
    def keyword_search_emails(self, query: str, max_results: int = 10, days_back: int = 30):
        """Fallback keyword search for emails"""
        try:
            start_date = datetime.now() - timedelta(days=days_back)
            results = []
            
            messages = self.inbox.Items
            keywords = query.lower().split()
            
            for message in messages:
                try:
                    if message.ReceivedTime >= start_date:
                        email_text = f"{message.Subject} {message.Body}".lower()
                        
                        # Calculate keyword match score
                        matches = sum(1 for keyword in keywords if keyword in email_text)
                        if matches > 0:
                            results.append({
                                'subject': message.Subject,
                                'sender': message.SenderEmailAddress,
                                'received_time': message.ReceivedTime,
                                'similarity_score': matches / len(keywords),
                                'entry_id': message.EntryID
                            })
                
                except Exception as e:
                    continue
            
            # Sort by relevance and return top results
            results.sort(key=lambda x: x['similarity_score'], reverse=True)
            return results[:max_results]
            
        except Exception as e:
            logger.error(f"Error in keyword search: {e}")
            return []


class OutlookAutomationGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Smart Outlook Email Automation Tool")
        self.root.geometry("800x600")
        
        self.automation_tool = OutlookAutomationTool()
        self.setup_gui()
        
    def setup_gui(self):
        """Setup the GUI interface"""
        # Create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Template Emails
        self.setup_template_tab(notebook)
        
        # Tab 2: Response Tracking
        self.setup_tracking_tab(notebook)
        
        # Tab 3: Attachment Management
        self.setup_attachment_tab(notebook)
        
        # Tab 4: Live Tracking
        self.setup_live_tracking_tab(notebook)
        
        # Tab 5: Email Search
        self.setup_search_tab(notebook)
    
    def setup_template_tab(self, notebook):
        """Setup template email tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="Template Emails")
        
        # Template file selection
        ttk.Label(frame, text="Email Template File:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.template_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.template_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse", command=self.browse_template_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Recipients CSV file
        ttk.Label(frame, text="Recipients CSV File:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.recipients_path = tk.StringVar()
        ttk.Entry(frame, textvariable=self.recipients_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse", command=self.browse_recipients_file).grid(row=1, column=2, padx=5, pady=5)
        
        # Subject template
        ttk.Label(frame, text="Subject Template:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.subject_template = tk.StringVar()
        ttk.Entry(frame, textvariable=self.subject_template, width=50).grid(row=2, column=1, columnspan=2, padx=5, pady=5)
        
        # Attachments
        ttk.Label(frame, text="Attachments (optional):").grid(row=3, column=0, sticky='w', padx=5, pady=5)
        self.attachments_listbox = tk.Listbox(frame, height=3)
        self.attachments_listbox.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(frame, text="Add Files", command=self.add_attachments).grid(row=3, column=2, padx=5, pady=5)
        
        # Send button
        ttk.Button(frame, text="Send Template Emails", command=self.send_template_emails).grid(row=4, column=0, columnspan=3, pady=20)
        
        # Results display
        self.template_results = scrolledtext.ScrolledText(frame, height=10, width=80)
        self.template_results.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
    
    def setup_tracking_tab(self, notebook):
        """Setup response tracking tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="Response Tracking")
        
        # Filters
        ttk.Label(frame, text="Subject Filter:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.tracking_subject_filter = tk.StringVar()
        ttk.Entry(frame, textvariable=self.tracking_subject_filter, width=30).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Days Back:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.tracking_days = tk.StringVar(value="7")
        ttk.Entry(frame, textvariable=self.tracking_days, width=10).grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Button(frame, text="Track Responses", command=self.track_responses).grid(row=0, column=4, padx=5, pady=5)
        
        # Results tree
        columns = ('Subject', 'Sender', 'Received', 'Attachments', 'Voting Response')
        self.tracking_tree = ttk.Treeview(frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.tracking_tree.heading(col, text=col)
            self.tracking_tree.column(col, width=150)
        
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=self.tracking_tree.yview)
        self.tracking_tree.configure(yscrollcommand=scrollbar.set)
        
        self.tracking_tree.grid(row=1, column=0, columnspan=5, padx=5, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=5, sticky='ns')
        
        # Export button
        ttk.Button(frame, text="Export to Excel", command=self.export_tracking_results).grid(row=2, column=0, columnspan=6, pady=10)
    
    def setup_attachment_tab(self, notebook):
        """Setup attachment management tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="Attachment Management")
        
        # Subject filter
        ttk.Label(frame, text="Subject Filter:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.attachment_subject_filter = tk.StringVar()
        ttk.Entry(frame, textvariable=self.attachment_subject_filter, width=40).grid(row=0, column=1, padx=5, pady=5)
        
        # Download folder
        ttk.Label(frame, text="Download Folder:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.download_folder = tk.StringVar()
        ttk.Entry(frame, textvariable=self.download_folder, width=40).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse", command=self.browse_download_folder).grid(row=1, column=2, padx=5, pady=5)
        
        # Merge Excel option
        self.merge_excel_var = tk.BooleanVar()
        ttk.Checkbutton(frame, text="Merge Excel Files", variable=self.merge_excel_var).grid(row=2, column=0, columnspan=2, sticky='w', padx=5, pady=5)
        
        # Download button
        ttk.Button(frame, text="Download Attachments", command=self.download_attachments).grid(row=3, column=0, columnspan=3, pady=10)
        
        # Results display
        self.attachment_results = scrolledtext.ScrolledText(frame, height=15, width=80)
        self.attachment_results.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
    
    def setup_live_tracking_tab(self, notebook):
        """Setup live tracking tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="Live Tracking")
        
        # Controls
        control_frame = ttk.Frame(frame)
        control_frame.pack(pady=10)
        
        ttk.Label(control_frame, text="Check Interval (minutes):").grid(row=0, column=0, padx=5)
        self.tracking_interval = tk.StringVar(value="5")
        ttk.Entry(control_frame, textvariable=self.tracking_interval, width=10).grid(row=0, column=1, padx=5)
        
        self.start_tracking_btn = ttk.Button(control_frame, text="Start Tracking", command=self.start_live_tracking)
        self.start_tracking_btn.grid(row=0, column=2, padx=5)
        
        self.stop_tracking_btn = ttk.Button(control_frame, text="Stop Tracking", command=self.stop_live_tracking, state='disabled')
        self.stop_tracking_btn.grid(row=0, column=3, padx=5)
        
        # Status
        self.tracking_status = tk.StringVar(value="Tracking stopped")
        ttk.Label(control_frame, textvariable=self.tracking_status).grid(row=1, column=0, columnspan=4, pady=5)
        
        # Live results
        self.live_results = scrolledtext.ScrolledText(frame, height=20, width=80)
        self.live_results.pack(padx=5, pady=5, fill='both', expand=True)
    
    def setup_search_tab(self, notebook):
        """Setup email search tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="Email Search")
        
        # Search controls
        search_frame = ttk.Frame(frame)
        search_frame.pack(pady=10)
        
        ttk.Label(search_frame, text="Search Query:").grid(row=0, column=0, sticky='w', padx=5)
        self.search_query = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.search_query, width=40).grid(row=0, column=1, padx=5)
        
        ttk.Label(search_frame, text="Max Results:").grid(row=0, column=2, padx=5)
        self.max_results = tk.StringVar(value="10")
        ttk.Entry(search_frame, textvariable=self.max_results, width=10).grid(row=0, column=3, padx=5)
        
        ttk.Label(search_frame, text="Days Back:").grid(row=1, column=0, padx=5)
        self.search_days = tk.StringVar(value="30")
        ttk.Entry(search_frame, textvariable=self.search_days, width=10).grid(row=1, column=1, padx=5)
        
        ttk.Button(search_frame, text="Search", command=self.search_emails).grid(row=1, column=2, columnspan=2, padx=5, pady=5)
        
        # Search results
        search_columns = ('Subject', 'Sender', 'Received', 'Relevance Score')
        self.search_tree = ttk.Treeview(frame, columns=search_columns, show='headings', height=15)
        
        for col in search_columns:
            self.search_tree.heading(col, text=col)
            self.search_tree.column(col, width=150)
        
        search_scrollbar = ttk.Scrollbar(frame, orient='vertical', command=self.search_tree.yview)
        self.search_tree.configure(yscrollcommand=search_scrollbar.set)
        
        self.search_tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        search_scrollbar.pack(side='right', fill='y')
    
    # Event handlers
    def browse_template_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if filename:
            self.template_path.set(filename)
    
    def browse_recipients_file(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if filename:
            self.recipients_path.set(filename)
    
    def add_attachments(self):
        filenames = filedialog.askopenfilenames()
        for filename in filenames:
            self.attachments_listbox.insert(tk.END, filename)
    
    def browse_download_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.download_folder.set(folder)
    
    def send_template_emails(self):
        try:
            # Validate inputs
            if not self.template_path.get() or not self.recipients_path.get():
                messagebox.showerror("Error", "Please select template and recipients files")
                return
            
            # Read recipients data
            recipients_df = pd.read_csv(self.recipients_path.get())
            recipients_data = recipients_df.to_dict('records')
            
            # Get attachments
            attachments = list(self.attachments_listbox.get(0, tk.END))
            
            # Send emails
            result = self.automation_tool.send_template_emails(
                self.template_path.get(),
                recipients_data,
                self.subject_template.get(),
                attachments if attachments else None
            )
            
            # Display results
            result_text = f"""
Email Sending Results:
=====================
Total Recipients: {result['total_recipients']}
Successfully Sent: {result['sent_count']}
Failed: {len(result['failed_recipients'])}

Failed Recipients:
{chr(10).join(result['failed_recipients'])}
            """
            
            self.template_results.delete(1.0, tk.END)
            self.template_results.insert(tk.END, result_text)
            
            messagebox.showinfo("Success", f"Successfully sent {result['sent_count']} emails!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails: {str(e)}")
    
    def track_responses(self):
        try:
            # Clear previous results
            for item in self.tracking_tree.get_children():
                self.tracking_tree.delete(item)
            
            # Get responses
            subject_filter = self.tracking_subject_filter.get() if self.tracking_subject_filter.get() else None
            days_back = int(self.tracking_days.get()) if self.tracking_days.get() else 7
            
            responses = self.automation_tool.track_email_responses(subject_filter, days_back)
            
            # Populate tree
            for response in responses:
                self.tracking_tree.insert('', 'end', values=(
                    response['subject'][:50] + "..." if len(response['subject']) > 50 else response['subject'],
                    response['sender_name'],
                    response['received_time'].strftime('%Y-%m-%d %H:%M'),
                    f"{response['attachment_count']} files" if response['has_attachments'] else "None",
                    response['voting_response']
                ))
            
            messagebox.showinfo("Success", f"Found {len(responses)} email responses")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to track responses: {str(e)}")
    
    def export_tracking_results(self):
        try:
            # Get data from tree
            data = []
            for item in self.tracking_tree.get_children():
                values = self.tracking_tree.item(item)['values']
                data.append(values)
            
            if not data:
                messagebox.showwarning("Warning", "No data to export")
                return
            
            # Create DataFrame and export
            df = pd.DataFrame(data, columns=['Subject', 'Sender', 'Received', 'Attachments', 'Voting Response'])
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if filename:
                df.to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Data exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def download_attachments(self):
        try:
            if not self.attachment_subject_filter.get():
                messagebox.showerror("Error", "Please enter a subject filter")
                return
            
            if not self.download_folder.get():
                messagebox.showerror("Error", "Please select a download folder")
                return
            
            # Download attachments
            result = self.automation_tool.download_attachments_by_subject(
                self.attachment_subject_filter.get(),
                self.download_folder.get(),
                self.merge_excel_var.get()
            )
            
            # Display results
            result_text = f"""
Attachment Download Results:
===========================
Total Files Downloaded: {result['total_downloads']}
Excel Files Found: {len(result['excel_files'])}
Merged File Created: {'Yes' if result['merged_file'] else 'No'}

Downloaded Files:
{chr(10).join([os.path.basename(f) for f in result['downloaded_files']])}

{f"Merged File: {result['merged_file']}" if result['merged_file'] else ""}
            """
            
            self.attachment_results.delete(1.0, tk.END)
            self.attachment_results.insert(tk.END, result_text)
            
            messagebox.showinfo("Success", f"Downloaded {result['total_downloads']} files!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download attachments: {str(e)}")
    
    def start_live_tracking(self):
        try:
            interval = int(self.tracking_interval.get())
            
            def new_email_callback(new_emails):
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                self.live_results.insert(tk.END, f"\n[{timestamp}] Found {len(new_emails)} new emails:\n")
                
                for email in new_emails:
                    self.live_results.insert(tk.END, f"  • {email['subject'][:60]}... from {email['sender']}\n")
                
                self.live_results.see(tk.END)
            
            self.automation_tool.start_live_tracking(new_email_callback, interval)
            
            self.start_tracking_btn.config(state='disabled')
            self.stop_tracking_btn.config(state='normal')
            self.tracking_status.set(f"Live tracking active (checking every {interval} minutes)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start live tracking: {str(e)}")
    
    def stop_live_tracking(self):
        try:
            self.automation_tool.stop_live_tracking()
            
            self.start_tracking_btn.config(state='normal')
            self.stop_tracking_btn.config(state='disabled')
            self.tracking_status.set("Tracking stopped")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to stop live tracking: {str(e)}")
    
    def search_emails(self):
        try:
            # Clear previous results
            for item in self.search_tree.get_children():
                self.search_tree.delete(item)
            
            if not self.search_query.get():
                messagebox.showerror("Error", "Please enter a search query")
                return
            
            # Perform search
            max_results = int(self.max_results.get()) if self.max_results.get() else 10
            days_back = int(self.search_days.get()) if self.search_days.get() else 30
            
            results = self.automation_tool.semantic_search_emails(
                self.search_query.get(),
                max_results,
                days_back
            )
            
            # Populate results tree
            for result in results:
                self.search_tree.insert('', 'end', values=(
                    result['subject'][:50] + "..." if len(result['subject']) > 50 else result['subject'],
                    result['sender'],
                    result['received_time'].strftime('%Y-%m-%d %H:%M'),
                    f"{result['similarity_score']:.3f}"
                ))
            
            messagebox.showinfo("Success", f"Found {len(results)} matching emails")
            
        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")
    
    def run(self):
        """Start the GUI application"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.automation_tool.stop_live_tracking()
            self.root.quit()


# Additional utility functions and configuration

def create_sample_template():
    """Create a sample email template"""
    template_content = """
Dear {name},

I hope this email finds you well. 

We are reaching out regarding {subject_matter}. Your participation in {project_name} is highly valued.

Please find the attached {document_type} for your review. We would appreciate your response by {deadline}.

Key details:
- Project: {project_name}
- Due Date: {deadline}
- Contact: {contact_person}

If you have any questions, please don't hesitate to reach out.

Best regards,
{sender_name}
{sender_title}
{company_name}
    """
    
    with open("sample_template.txt", "w") as f:
        f.write(template_content)
    
    print("Sample template created: sample_template.txt")

def create_sample_recipients_csv():
    """Create a sample recipients CSV file"""
    sample_data = {
        'name': ['John Doe', 'Jane Smith', 'Mike Johnson', 'Sarah Wilson'],
        'email': ['john.doe@example.com', 'jane.smith@example.com', 'mike.johnson@example.com', 'sarah.wilson@example.com'],
        'project_name': ['Project Alpha', 'Project Beta', 'Project Alpha', 'Project Gamma'],
        'deadline': ['2024-01-15', '2024-01-20', '2024-01-15', '2024-01-25'],
        'contact_person': ['Alice Brown', 'Bob Davis', 'Alice Brown', 'Carol White'],
        'subject_matter': ['Budget Review', 'Technical Assessment', 'Budget Review', 'Quality Audit'],
        'document_type': ['spreadsheet', 'technical document', 'spreadsheet', 'checklist']
    }
    
    df = pd.DataFrame(sample_data)
    df.to_csv("sample_recipients.csv", index=False)
    print("Sample recipients CSV created: sample_recipients.csv")

def install_requirements():
    """Install required packages"""
    requirements = [
        'pywin32',
        'pandas',
        'sentence-transformers',
        'scikit-learn',
        'numpy',
        'openpyxl'
    ]
    
    import subprocess
    import sys
    
    for package in requirements:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Successfully installed {package}")
        except subprocess.CalledProcessError:
            print(f"Failed to install {package}")

class EmailTemplateBuilder:
    """Helper class for building email templates with GUI"""
    
    def __init__(self):
        self.template_window = None
    
    def open_template_builder(self):
        """Open template builder window"""
        self.template_window = tk.Toplevel()
        self.template_window.title("Email Template Builder")
        self.template_window.geometry("600x500")
        
        # Template text area
        ttk.Label(self.template_window, text="Email Template:").pack(pady=5)
        self.template_text = scrolledtext.ScrolledText(self.template_window, height=15, width=70)
        self.template_text.pack(padx=10, pady=5)
        
        # Variables list
        ttk.Label(self.template_window, text="Available Variables (use as {variable_name}):").pack(pady=5)
        variables_frame = ttk.Frame(self.template_window)
        variables_frame.pack(pady=5)
        
        common_vars = [
            'name', 'email', 'company', 'project_name', 'deadline', 
            'contact_person', 'phone', 'address', 'subject_matter'
        ]
        
        for i, var in enumerate(common_vars):
            ttk.Button(variables_frame, text=f"{{{var}}}", 
                      command=lambda v=var: self.insert_variable(v)).grid(row=i//3, column=i%3, padx=2, pady=2)
        
        # Save button
        ttk.Button(self.template_window, text="Save Template", 
                  command=self.save_template).pack(pady=10)
    
    def insert_variable(self, var_name):
        """Insert variable into template"""
        self.template_text.insert(tk.INSERT, f"{{{var_name}}}")
    
    def save_template(self):
        """Save template to file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            with open(filename, 'w') as f:
                f.write(self.template_text.get(1.0, tk.END))
            messagebox.showinfo("Success", f"Template saved to {filename}")

class EmailAnalytics:
    """Email analytics and reporting functionality"""
    
    def __init__(self, automation_tool):
        self.automation_tool = automation_tool
    
    def generate_response_analytics(self, subject_filter=None, days_back=30):
        """Generate analytics report for email responses"""
        try:
            responses = self.automation_tool.track_email_responses(subject_filter, days_back)
            
            if not responses:
                return {"error": "No responses found"}
            
            # Calculate statistics
            total_responses = len(responses)
            responses_with_attachments = sum(1 for r in responses if r['has_attachments'])
            voting_responses = [r for r in responses if r['is_voting_response']]
            
            # Voting analysis
            voting_stats = {}
            for response in voting_responses:
                vote = response['voting_response']
                voting_stats[vote] = voting_stats.get(vote, 0) + 1
            
            # Response time analysis
            response_times = []
            for response in responses:
                # Assuming we can match with sent emails (simplified)
                response_times.append(response['received_time'])
            
            # Daily response distribution
            daily_responses = {}
            for response in responses:
                date_key = response['received_time'].strftime('%Y-%m-%d')
                daily_responses[date_key] = daily_responses.get(date_key, 0) + 1
            
            analytics = {
                'total_responses': total_responses,
                'responses_with_attachments': responses_with_attachments,
                'attachment_percentage': (responses_with_attachments / total_responses) * 100,
                'voting_responses': len(voting_responses),
                'voting_stats': voting_stats,
                'daily_responses': daily_responses,
                'top_responders': self.get_top_responders(responses),
                'response_rate_by_day': daily_responses
            }
            
            return analytics
            
        except Exception as e:
            logger.error(f"Error generating analytics: {e}")
            return {"error": str(e)}
    
    def get_top_responders(self, responses, limit=10):
        """Get top email responders"""
        responder_counts = {}
        for response in responses:
            email = response['sender']
            responder_counts[email] = responder_counts.get(email, 0) + 1
        
        # Sort by count and return top responders
        sorted_responders = sorted(responder_counts.items(), key=lambda x: x[1], reverse=True)
        return sorted_responders[:limit]
    
    def export_analytics_report(self, analytics, filename):
        """Export analytics to Excel report"""
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Summary sheet
                summary_data = {
                    'Metric': ['Total Responses', 'Responses with Attachments', 'Attachment Percentage', 'Voting Responses'],
                    'Value': [
                        analytics['total_responses'],
                        analytics['responses_with_attachments'],
                        f"{analytics['attachment_percentage']:.1f}%",
                        analytics['voting_responses']
                    ]
                }
                pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
                
                # Voting stats
                if analytics['voting_stats']:
                    voting_df = pd.DataFrame(list(analytics['voting_stats'].items()), 
                                           columns=['Vote Type', 'Count'])
                    voting_df.to_excel(writer, sheet_name='Voting Analysis', index=False)
                
                # Daily responses
                if analytics['daily_responses']:
                    daily_df = pd.DataFrame(list(analytics['daily_responses'].items()), 
                                          columns=['Date', 'Response Count'])
                    daily_df.to_excel(writer, sheet_name='Daily Responses', index=False)
                
                # Top responders
                if analytics['top_responders']:
                    responders_df = pd.DataFrame(analytics['top_responders'], 
                                               columns=['Email', 'Response Count'])
                    responders_df.to_excel(writer, sheet_name='Top Responders', index=False)
            
            return True
            
        except Exception as e:
            logger.error(f"Error exporting analytics: {e}")
            return False


# Main execution
if __name__ == "__main__":
    try:
        # Check if running in development mode
        import sys
        if len(sys.argv) > 1 and sys.argv[1] == "--setup":
            print("Setting up sample files...")
            create_sample_template()
            create_sample_recipients_csv()
            print("Setup complete!")
            sys.exit(0)
        
        if len(sys.argv) > 1 and sys.argv[1] == "--install":
            print("Installing required packages...")
            install_requirements()
            print("Installation complete!")
            sys.exit(0)
        
        # Launch GUI application
        print("Starting Smart Outlook Email Automation Tool...")
        app = OutlookAutomationGUI()
        app.run()
        
    except ImportError as e:
        print(f"Missing required package: {e}")
        print("Run 'python outlook_automation.py --install' to install required packages")
    except Exception as e:
        print(f"Error starting application: {e}")
        input("Press Enter to exit...")
