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

class OCRTextVisibilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR Text Visibility Control")
        self.root.geometry("1000x600")
        
        # Setup logging and speech
        self.logger = self.setup_logging()
        self.speaker = self.setup_speech()
        
        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create control buttons
        self.create_control_buttons()
        
        # Create text display area
        self.create_text_display()
        
        # Initialize variables
        self.current_image = None
        self.detected_text = ""
        
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
        """Create the control buttons panel"""
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        # Select Image button
        self.select_btn = ttk.Button(
            button_frame,
            text="Select Image",
            command=self.select_image
        )
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        # Process Image button
        self.process_btn = ttk.Button(
            button_frame,
            text="Process Image",
            command=self.process_image,
            state=tk.DISABLED
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        # Speak Text button
        self.speak_btn = ttk.Button(
            button_frame,
            text="Speak Text",
            command=self.speak_detected_text,
            state=tk.DISABLED
        )
        self.speak_btn.pack(side=tk.LEFT, padx=5)
        
        # Show/Hide Text checkbox
        self.visibility_var = tk.BooleanVar(value=True)
        self.visibility_check = ttk.Checkbutton(
            button_frame,
            text="Show/Hide Text",
            variable=self.visibility_var,
            command=self.toggle_text_visibility
        )
        self.visibility_check.pack(side=tk.LEFT, padx=5)
    
    def create_text_display(self):
        """Create the text display area"""
        # Text display
        self.text_display = tk.Text(
            self.main_frame,
            wrap=tk.WORD,
            height=10,
            width=50
        )
        self.text_display.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def select_image(self):
        """Handle image selection"""
        file_path = askopenfilename(
    title="Select Image File",
    filetypes=[("All Files", "*.*")]
)

        
        if file_path:
            try:
                self.current_image = cv2.imread(file_path)
                if self.current_image is None:
                    raise ValueError("Failed to load image")
                
                self.process_btn.config(state=tk.NORMAL)
                self.logger.info("Image loaded successfully")
                
            except Exception as e:
                error_message = f"Error loading image: {str(e)}"
                self.logger.error(error_message)
                messagebox.showerror("Error", error_message)
    
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
            self.text_display.delete(1.0, tk.END)
            self.text_display.insert(tk.END, self.detected_text)
            
            # Enable speak button
            self.speak_btn.config(state=tk.NORMAL)
            
            # Display image
            cv2.imshow("Processed Image", self.current_image)
            
            self.logger.info("Image processed successfully")
            
        except Exception as e:
            error_message = f"Error processing image: {str(e)}"
            self.logger.error(error_message)
            messagebox.showerror("Error", error_message)
    
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
            self.text_display.pack(fill=tk.BOTH, expand=True, pady=5)
        else:
            self.text_display.pack_forget()

def main():
    root = tk.Tk()
    app = OCRTextVisibilityApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
