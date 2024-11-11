import cv2
import numpy as np
import requests
import io
import json
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename
import win32com.client
import logging
import subprocess
import sys

def setup_logging():
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

def setup_speech():
    """Initialize the Windows SAPI speech engine."""
    try:
        # Check if pywin32 is installed
        try:
            import win32com.client
        except ImportError:
            print("Installing required package: pywin32...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
            import win32com.client
        
        # Initialize speech engine
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        return speaker
    except Exception as e:
        print(f"Error setting up speech: {str(e)}")
        return None

def speak_text(speaker, text):
    """Speak the given text using Windows SAPI."""
    try:
        if speaker and text:
            speaker.Speak(text)
    except Exception as e:
        print(f"Error speaking text: {str(e)}")

# Set up logging
logger = setup_logging()

# Initialize speech engine
speaker = setup_speech()

# Create a Tkinter window (hidden)
Tk().withdraw()

# Open file dialog for user to select image file
file_path = askopenfilename(
    title="Select Image File", 
    filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp")]
)

if file_path:
    try:
        # Read selected image
        img = cv2.imread(file_path)
        if img is None:
            raise ValueError("Failed to load image")
        
        height, width, _ = img.shape
        # Use full image as ROI
        roi = img
        
        # Perform OCR
        url_api = "https://api.ocr.space/parse/image"
        _, compressedimage = cv2.imencode(".jpg", roi, [1, 90])
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
        text_detected = parsed_results.get("ParsedText")
        
        # Display detected text
        print("\nDetected Text:")
        print("-" * 50)
        print(text_detected)
        print("-" * 50)
        
        # Speak the detected text
        if text_detected:
            logger.info("Speaking detected text...")
            speak_text(speaker, text_detected)
        else:
            logger.warning("No text detected in image")
            speak_text(speaker, "No text detected in image")
        
        # Display images
        cv2.imshow("Region of Interest", roi)
        cv2.imshow("Original Image", img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
        
    except Exception as e:
        error_message = f"Error processing image: {str(e)}"
        logger.error(error_message)
        messagebox.showerror("Error", error_message)
        speak_text(speaker, "An error occurred while processing the image")
        
else:
    message = "No file selected."
    print(message)
    logger.info(message)
    speak_text(speaker, message)
