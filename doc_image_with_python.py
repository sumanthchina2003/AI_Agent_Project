import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename
import cv2
import numpy as np
import requests
import io
import json
import win32com.client
import logging
import subprocess
import sys
import PyPDF2
from pdf2image import convert_from_path
import pytesseract
import os
import tempfile

class OCRTextVisibilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR, PDF Text Visibility Control")
        self.root.geometry("1000x800")
        
        # Setup logging and speech
        self.logger = self.setup_logging()
        self.speaker = self.setup_speech()
        
        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Initialize variables
        self.current_image = None
        self.detected_text = ""
        self.current_pdf = None
        self.pdf_pages = []
        self.current_page = 0
        
        # Create GUI elements
        self.create_control_buttons()
        self.create_pdf_controls()
        self.create_text_display()
        
    def setup_logging(self):
        """Configure logging for the application."""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('ocr_reader.log'),
                logging.StreamHandler()
            ]
        )
        return logging.getLogger(__name__)
    
    def setup_speech(self):
        """Initialize the Windows SAPI speech engine."""
        try:
            try:
                import win32com.client
            except ImportError:
                print("Installing required package: pywin32...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
                import win32com.client
            
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            return speaker
        except Exception as e:
            self.logger.error(f"Error setting up speech: {str(e)}")
            return None
    
    def create_control_buttons(self):
        """Create the main control buttons panel"""
        button_frame = ttk.LabelFrame(self.main_frame, text="File Controls")
        button_frame.pack(fill=tk.X, pady=5)
        
        # Select Image button
        self.select_img_btn = ttk.Button(
            button_frame,
            text="Select Image",
            command=self.select_image
        )
        self.select_img_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Select PDF button
        self.select_pdf_btn = ttk.Button(
            button_frame,
            text="Select PDF",
            command=self.select_pdf
        )
        self.select_pdf_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Process button
        self.process_btn = ttk.Button(
            button_frame,
            text="Process File",
            command=self.process_file,
            state=tk.DISABLED
        )
        self.process_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Speak Text button
        self.speak_btn = ttk.Button(
            button_frame,
            text="Speak Text",
            command=self.speak_detected_text,
            state=tk.DISABLED
        )
        self.speak_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Show/Hide Text checkbox
        self.visibility_var = tk.BooleanVar(value=True)
        self.visibility_check = ttk.Checkbutton(
            button_frame,
            text="Show/Hide Text",
            variable=self.visibility_var,
            command=self.toggle_text_visibility
        )
        self.visibility_check.pack(side=tk.LEFT, padx=5, pady=5)
    
    def create_pdf_controls(self):
        """Create PDF navigation controls"""
        self.pdf_frame = ttk.LabelFrame(self.main_frame, text="PDF Navigation")
        self.pdf_frame.pack(fill=tk.X, pady=5)
        
        # Previous page button
        self.prev_btn = ttk.Button(
            self.pdf_frame,
            text="Previous Page",
            command=self.previous_page,
            state=tk.DISABLED
        )
        self.prev_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Page counter label
        self.page_label = ttk.Label(self.pdf_frame, text="Page: 0/0")
        self.page_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Next page button
        self.next_btn = ttk.Button(
            self.pdf_frame,
            text="Next Page",
            command=self.next_page,
            state=tk.DISABLED
        )
        self.next_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Process current page button
        self.process_page_btn = ttk.Button(
            self.pdf_frame,
            text="Process Current Page",
            command=self.process_current_page,
            state=tk.DISABLED
        )
        self.process_page_btn.pack(side=tk.LEFT, padx=5, pady=5)
    
    def create_text_display(self):
        """Create the text display area"""
        # Text display with scrollbar
        text_frame = ttk.Frame(self.main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.text_scroll = ttk.Scrollbar(text_frame)
        self.text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.text_display = tk.Text(
            text_frame,
            wrap=tk.WORD,
            height=20,
            width=70,
            yscrollcommand=self.text_scroll.set
        )
        self.text_display.pack(fill=tk.BOTH, expand=True)
        self.text_scroll.config(command=self.text_display.yview)
    
    def select_image(self):
        """Handle image selection"""
        file_path = askopenfilename(
            title="Select Image File",
            filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp")]
        )
        
        if file_path:
            try:
                self.current_image = cv2.imread(file_path)
                if self.current_image is None:
                    raise ValueError("Failed to load image")
                
                self.current_pdf = None
                self.pdf_pages = []
                self.current_page = 0
                self.process_btn.config(state=tk.NORMAL)
                self.logger.info("Image loaded successfully")
                
            except Exception as e:
                error_message = f"Error loading image: {str(e)}"
                self.logger.error(error_message)
                messagebox.showerror("Error", error_message)
    
    def select_pdf(self):
        """Handle PDF selection"""
        file_path = askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF Files", "*.pdf")]
        )
        
        if file_path:
            try:
                # Convert PDF to images
                self.pdf_pages = convert_from_path(file_path)
                self.current_page = 0
                self.current_pdf = file_path
                self.current_image = None
                
                # Update UI
                self.process_btn.config(state=tk.NORMAL)
                self.prev_btn.config(state=tk.NORMAL)
                self.next_btn.config(state=tk.NORMAL)
                self.process_page_btn.config(state=tk.NORMAL)
                self.page_label.config(text=f"Page: 1/{len(self.pdf_pages)}")
                
                self.logger.info(f"PDF loaded successfully: {len(self.pdf_pages)} pages")
                
            except Exception as e:
                error_message = f"Error loading PDF: {str(e)}"
                self.logger.error(error_message)
                messagebox.showerror("Error", error_message)
    
    def process_file(self):
        """Process the selected file (image or PDF)"""
        if self.current_image is not None:
            self.process_image()
        elif self.current_pdf is not None:
            self.process_current_page()
    
    def process_image(self):
        """Process the selected image with OCR"""
        if self.current_image is None:
            return
        
        try:
            # Perform OCR
            url_api = "https://api.ocr.space/parse/image"
            _, compressedimage = cv2.imencode(".jpg", self.current_image, [1, 90])
            file_bytes = io.BytesIO(compressedimage)
            
            # Send OCR request
            result = requests.post(
                url_api,
                files={"screenshot.jpg": file_bytes},
                data={
                    "apikey": "helloworld",
                    "language": "eng"
                }
            )
            
            # Parse results
            result = result.content.decode()
            result = json.loads(result)
            parsed_results = result.get("ParsedResults")[0]
            self.detected_text = parsed_results.get("ParsedText")
            
            # Update display
            self.update_text_display(self.detected_text)
            self.speak_btn.config(state=tk.NORMAL)
            
            # Display image
            cv2.imshow("Processed Image", self.current_image)
            
            self.logger.info("Image processed successfully")
            
        except Exception as e:
            error_message = f"Error processing image: {str(e)}"
            self.logger.error(error_message)
            messagebox.showerror("Error", error_message)
    
    def process_current_page(self):
        """Process the current PDF page"""
        if not self.pdf_pages:
            return
        
        try:
            # Convert current page to image array
            current_page_image = np.array(self.pdf_pages[self.current_page])
            
            # Perform OCR using pytesseract
            self.detected_text = pytesseract.image_to_string(current_page_image)
            
            # Update display
            self.update_text_display(self.detected_text)
            self.speak_btn.config(state=tk.NORMAL)
            
            # Display image
            cv2.imshow("Current PDF Page", cv2.cvtColor(current_page_image, cv2.COLOR_RGB2BGR))
            
            self.logger.info(f"PDF page {self.current_page + 1} processed successfully")
            
        except Exception as e:
            error_message = f"Error processing PDF page: {str(e)}"
            self.logger.error(error_message)
            messagebox.showerror("Error", error_message)
    
    def update_text_display(self, text):
        """Update the text display with new text"""
        self.text_display.delete(1.0, tk.END)
        self.text_display.insert(tk.END, text)
    
    def speak_detected_text(self):
        """Speak the detected text"""
        if self.detected_text and self.speaker:
            try:
                self.speaker.Speak(self.detected_text)
                self.logger.info("Speaking detected text")
            except Exception as e:
                error_message = f"Error speaking text: {str(e)}"
                self.logger.error(error_message)
                messagebox.showerror("Error", error_message)
    
    def toggle_text_visibility(self):
        """Toggle the visibility of the text display"""
        if self.visibility_var.get():
            self.text_display.pack(fill=tk.BOTH, expand=True)
            self.text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        else:
            self.text_display.pack_forget()
            self.text_scroll.pack_forget()
    
    def previous_page(self):
        """Go to previous PDF page"""
        if self.current_page > 0:
            self.current_page -= 1
            self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_pages)}")
    
    def next_page(self):
        """Go to next PDF page"""
        if self.current_page < len(self.pdf_pages) - 1:
            self.current_page += 1
            self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_pages)}")

def main():
    # Install required packages if not present
    required_packages = ['PyPDF2', 'pdf2image', 'pytesseract']
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Installing required package: {package}")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
    
    root = tk.Tk()
    app = OCRTextVisibilityApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
