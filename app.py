"""
Professional WhatsApp Bulk Sender - Flask Web Application
Complete with modern frontend, real-time progress, and file upload
"""

from flask import Flask, render_template, request, jsonify, send_file
from flask_socketio import SocketIO, emit
from flask_cors import CORS
import pandas as pd
import time
import webbrowser
import pyautogui
import urllib.parse
import os
from datetime import datetime
from werkzeug.utils import secure_filename
import threading
import logging
import traceback

app = Flask(__name__)
app.config['SECRET_KEY'] = 'whatsapp_bulk_sender_secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Enable CORS
CORS(app)

socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class WhatsAppBulkSender:
    def __init__(self, excel_path, default_country_code="+91", socketio_instance=None):
        self.excel_path = excel_path
        self.default_country_code = default_country_code
        self.socketio = socketio_instance
        
        self.open_delay = 10
        self.send_delay = 5
        self.image_delay = 3
        self.typing_delay = 0.1
        
        self.success_count = 0
        self.failure_count = 0
        self.failed_numbers = []
        self.is_running = False
        self.is_paused = False
        
    def emit_log(self, message, level='info'):
        """Emit log message to frontend"""
        if self.socketio:
            self.socketio.emit('log_message', {
                'message': message,
                'level': level,
                'timestamp': datetime.now().strftime('%H:%M:%S')
            })
        logger.info(f"[{level.upper()}] {message}")
    
    def emit_progress(self, current, total):
        """Emit progress update to frontend"""
        if self.socketio:
            percentage = int((current / total) * 100) if total > 0 else 0
            self.socketio.emit('progress_update', {
                'current': current,
                'total': total,
                'percentage': percentage,
                'success': self.success_count,
                'failed': self.failure_count
            })
    
    def format_phone_number(self, number):
        number = str(number).strip()
        cleaned = ''.join(c for c in number if c.isdigit() or c == '+')
        if not cleaned.startswith('+'):
            cleaned = self.default_country_code + cleaned
        return cleaned
    
    def prepare_message(self, message):
        message = str(message).strip()
        message = message.replace("\n", "%0A")
        message = urllib.parse.quote(message, safe='%')
        return message
    
    def send_message(self, number, message):
        try:
            url = f'whatsapp://send?phone={number}&text={message}'
            self.emit_log(f"Opening chat for {number}...")
            webbrowser.open(url)
            time.sleep(self.open_delay)
            pyautogui.press('enter')
            self.emit_log(f"✅ Message sent to {number}", 'success')
            time.sleep(self.send_delay)
            return True
        except Exception as e:
            self.emit_log(f"❌ Failed to send to {number}: {e}", 'error')
            return False
    
    def send_image(self, number, image_path):
        if not image_path or pd.isna(image_path):
            return True
        
        image_path = str(image_path).strip()
        if not os.path.exists(image_path):
            self.emit_log(f"⚠️ Image not found: {image_path}", 'warning')
            return False
        
        try:
            self.emit_log(f"📎 Attaching image for {number}...")
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'o')
            time.sleep(self.image_delay)
            pyautogui.write(image_path, interval=self.typing_delay)
            time.sleep(self.image_delay)
            pyautogui.press('enter')
            time.sleep(self.image_delay)
            pyautogui.press('enter')
            time.sleep(self.send_delay)
            self.emit_log(f"✅ Image sent to {number}", 'success')
            return True
        except Exception as e:
            self.emit_log(f"❌ Failed to send image: {e}", 'error')
            return False
    
    def run(self):
        self.is_running = True
        self.emit_log("🚀 Starting WhatsApp Bulk Sender...", 'info')
        
        try:
            df = pd.read_excel(self.excel_path)
            df.columns = df.columns.str.strip()
            
            required_columns = ["Number", "Message"]
            missing = [col for col in required_columns if col not in df.columns]
            if missing:
                self.emit_log(f"❌ Missing columns: {missing}", 'error')
                self.is_running = False
                return
            
            total = len(df)
            self.emit_log(f"📊 Found {total} messages to send", 'info')
            self.emit_log("⚠️ Do NOT use keyboard/mouse during execution!", 'warning')
            
            for i in range(5, 0, -1):
                self.emit_log(f"⏳ Starting in {i} seconds...", 'info')
                time.sleep(1)
            
            for index, row in df.iterrows():
                if not self.is_running:
                    self.emit_log("⏹️ Process stopped by user", 'warning')
                    break
                
                while self.is_paused:
                    time.sleep(0.5)
                
                try:
                    raw_number = row["Number"]
                    raw_message = row["Message"]
                    image_path = row.get("Image", "")
                    
                    number = self.format_phone_number(raw_number)
                    message = self.prepare_message(raw_message)
                    
                    self.emit_log(f"\n📱 [{index+1}/{total}] Processing {number}...", 'info')
                    
                    if self.send_message(number, message):
                        if image_path and not pd.isna(image_path):
                            self.send_image(number, image_path)
                        self.success_count += 1
                    else:
                        self.failure_count += 1
                        self.failed_numbers.append(number)
                    
                    self.emit_progress(index + 1, total)
                    
                except Exception as e:
                    self.emit_log(f"❌ Error processing row {index+1}: {e}", 'error')
                    self.failure_count += 1
                    self.failed_numbers.append(str(raw_number))
            
            self.emit_summary()
            
        except Exception as e:
            self.emit_log(f"❌ Fatal error: {e}", 'error')
            logger.error(traceback.format_exc())
        finally:
            self.is_running = False
            if self.socketio:
                self.socketio.emit('process_complete', {
                    'success': self.success_count,
                    'failed': self.failure_count,
                    'failed_numbers': self.failed_numbers
                })
    
    def emit_summary(self):
        self.emit_log("\n" + "="*50, 'info')
        self.emit_log("📊 EXECUTION SUMMARY", 'info')
        self.emit_log("="*50, 'info')
        self.emit_log(f"Total: {self.success_count + self.failure_count}", 'info')
        self.emit_log(f"✅ Successful: {self.success_count}", 'success')
        self.emit_log(f"❌ Failed: {self.failure_count}", 'error')
        
        if self.failed_numbers:
            self.emit_log("\n⚠️ Failed Numbers:", 'warning')
            for num in self.failed_numbers:
                self.emit_log(f"  • {num}", 'warning')
        
        self.emit_log("="*50, 'info')
        self.emit_log("✅ Process completed!", 'success')

# Global sender instance
current_sender = None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        logger.info("Upload request received")
        
        if 'file' not in request.files:
            logger.error("No file in request")
            return jsonify({'success': False, 'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        logger.info(f"File received: {file.filename}")
        
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            logger.error("Invalid file type")
            return jsonify({'success': False, 'error': 'Only Excel files (.xlsx, .xls) are allowed'}), 400
        
        # Save file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        logger.info(f"Saving file to: {filepath}")
        file.save(filepath)
        
        # Validate Excel file
        logger.info("Reading Excel file")
        df = pd.read_excel(filepath)
        df.columns = df.columns.str.strip()
        
        required_columns = ["Number", "Message"]
        missing = [col for col in required_columns if col not in df.columns]
        
        if missing:
            os.remove(filepath)
            error_msg = f'Missing required columns: {", ".join(missing)}'
            logger.error(error_msg)
            return jsonify({'success': False, 'error': error_msg}), 400
        
        # Prepare preview data
        preview_data = []
        for _, row in df.head(5).iterrows():
            preview_row = {}
            for col in df.columns:
                value = row[col]
                if pd.isna(value):
                    preview_row[col] = '-'
                else:
                    preview_row[col] = str(value)[:50]  # Limit preview length
            preview_data.append(preview_row)
        
        response_data = {
            'success': True,
            'filepath': filepath,
            'filename': file.filename,
            'total_rows': len(df),
            'columns': list(df.columns),
            'preview': preview_data
        }
        
        logger.info(f"Upload successful: {len(df)} rows")
        return jsonify(response_data), 200
    
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        logger.error(f"Upload error: {error_msg}")
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'error': error_msg}), 500

@socketio.on('connect')
def handle_connect():
    logger.info('Client connected')
    emit('connected', {'message': 'Connected to server'})

@socketio.on('disconnect')
def handle_disconnect():
    logger.info('Client disconnected')

@socketio.on('start_sending')
def handle_start_sending(data):
    global current_sender
    
    logger.info(f"Start sending request: {data}")
    
    if current_sender and current_sender.is_running:
        emit('error', {'message': 'A process is already running!'})
        return
    
    filepath = data.get('filepath')
    country_code = data.get('country_code', '+91')
    
    if not filepath or not os.path.exists(filepath):
        emit('error', {'message': 'Excel file not found!'})
        return
    
    current_sender = WhatsAppBulkSender(filepath, country_code, socketio)
    
    # Run in separate thread
    thread = threading.Thread(target=current_sender.run)
    thread.daemon = True
    thread.start()
    
    emit('process_started', {'message': 'Process started successfully!'})

@socketio.on('pause_sending')
def handle_pause():
    global current_sender
    if current_sender:
        current_sender.is_paused = not current_sender.is_paused
        status = 'paused' if current_sender.is_paused else 'resumed'
        emit('process_paused', {'status': status})

@socketio.on('stop_sending')
def handle_stop():
    global current_sender
    if current_sender:
        current_sender.is_running = False
        emit('process_stopped', {'message': 'Process stopped!'})

@app.route('/download-template')
def download_template():
    """Generate and download sample Excel template"""
    try:
        df = pd.DataFrame({
            'Number': ['9876543210', '9123456780'],
            'Message': ['Hello! This is a test message.', 'Welcome to WhatsApp Bulk Sender!'],
            'Image': ['C:\\path\\to\\image1.jpg', '']
        })
        
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template.xlsx')
        df.to_excel(template_path, index=False)
        
        return send_file(template_path, as_attachment=True, download_name='whatsapp_template.xlsx')
    except Exception as e:
        logger.error(f"Template download error: {e}")
        return jsonify({'error': str(e)}), 500

@app.errorhandler(413)
def too_large(e):
    return jsonify({'success': False, 'error': 'File too large. Maximum size is 16MB'}), 413

@app.errorhandler(500)
def internal_error(e):
    logger.error(f"Internal error: {e}")
    return jsonify({'success': False, 'error': 'Internal server error'}), 500

if __name__ == '__main__':
    print("""
    ╔═══════════════════════════════════════════════════════╗
    ║     WHATSAPP BULK SENDER PRO - WEB INTERFACE          ║
    ║              Server Starting...                       ║
    ╚═══════════════════════════════════════════════════════╝
    """)
    print("\n✅ Server running at: http://localhost:5000")
    print("📱 Open this URL in your browser to access the interface\n")
    
    socketio.run(app, debug=True, host='0.0.0.0', port=5000, allow_unsafe_werkzeug=True)